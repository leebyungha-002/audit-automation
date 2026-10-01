"""계정 대사. 회사 계정리스트와 지적사항의 관련 계정을 표준 계정 분류로 정규화해서 연결한다.

1차는 사전(account_taxonomy.yaml) 매칭, 2차는 신뢰도가 낮은 것만 Claude API로 의미 매칭한다.
외부로 보내는 것은 계정명·계정코드뿐이다 (회사명, 금액, 건수는 보내지 않는다).
"""
import json
import re
from pathlib import Path

import pandas as pd

from .db import Store, now
from .llm import LLM, LLMError

STATS_SHEET = "05_계정명 리스트_통계"
NAME_HEADERS = ("계정명", "계정과목", "계정과목명")
CODE_HEADERS = ("계정코드", "코드")
LLM_CONFIDENCE = {"high": 0.9, "medium": 0.7, "low": 0.5}
BATCH = 40

MAPPING_SYSTEM = """당신은 한국 기업의 계정과목명을 표준 계정 분류로 나누는 보조자입니다.
결과는 회계감사 실무자가 감리 지적사례와 회사 계정을 연결하는 데 쓰며, 사람이 다시 검토합니다.

입력은 번호가 붙은 계정과목명 목록입니다. 각 계정을 아래 표준 분류 중 가장 알맞은 하나에 배정하세요.
- 계정명 뒤의 괄호·밑줄·하이픈 부분은 보통 세부 구분(제조/판관, 거래처, 용도)입니다. 계정의 본래 성격으로 판단합니다.
- 알맞은 분류가 없거나 계정명만으로 성격을 알 수 없으면 억지로 맞추지 말고 category를 null로 둡니다(본지점 계정, 집합·대체 계정, 의미 없는 값 등).
- confidence는 계정명만으로 분류가 분명하면 high, 통상 그렇게 쓰이지만 회사마다 다를 수 있으면 medium, 추정이면 low.
- 입력에 있는 모든 번호에 대해 하나씩 답합니다.

표준 분류와 각 분류의 대표 키워드:
{categories}"""


def _norm(name: str) -> str:
    return re.sub(r"\s+", "", str(name)).lower()


class Taxonomy:
    def __init__(self, raw: dict):
        self.names = [c["name"] for c in raw["categories"]]
        self.overrides = {_norm(k): v for k, v in (raw.get("overrides") or {}).items()}
        self.ignore = {_norm(t) for t in raw.get("finding_term_ignore") or []}
        self._keywords = []  # (tier, 정규화 키워드, 분류) — tier 1 = strong
        for c in raw["categories"]:
            for tier, key in ((1, "strong_keywords"), (0, "keywords")):
                self._keywords += [(tier, _norm(k), c["name"]) for k in c.get(key) or []]
        self.described = "\n".join(
            f"- {c['name']}: {', '.join([*(c.get('strong_keywords') or []), *(c.get('keywords') or [])][:12])}"
            for c in raw["categories"])

    def match(self, name: str) -> tuple[str, float, str] | None:
        """(분류, 신뢰도, 근거 키워드). 맞는 키워드가 없으면 None."""
        norm = _norm(name)
        if norm in self.overrides:
            return self.overrides[norm], 1.0, "override"
        strong = [(len(kw), cat, kw) for tier, kw, cat in self._keywords if tier and kw in norm]
        if strong:
            _, cat, kw = max(strong)
            return cat, (1.0 if norm == kw else 0.9), kw

        # 괄호·하이픈·밑줄·슬래시 앞부분이 계정의 본래 이름이다. 우리말 복합명사는 끝말이 성격을
        # 정하므로("상품매출원가"는 원가, "반제품 매출"은 매출) 본래 이름에서 가장 뒤에서 끝나는 키워드를 고른다.
        core = re.split(r"[(\[\-_/:]", norm, maxsplit=1)[0] or norm
        for text, in_core in ((core, True), (norm, False)):
            found = []
            for _, kw, cat in self._keywords:
                pos = text.rfind(kw)
                if pos >= 0:
                    found.append((pos + len(kw), len(kw), cat, kw))
            if not found:
                continue
            end, _, cat, kw = max(found)
            if not in_core:
                return cat, 0.6, kw
            return cat, (1.0 if core == kw else 0.85 if end == len(core) else 0.7), kw
        return None


def load_account_list(path: Path) -> tuple[str | None, pd.DataFrame]:
    """계정리스트를 읽어 (회사명, DataFrame[account_name, account_code, n_debit, n_credit, n_vouchers])를 돌려준다.

    journal_analyzer 분석결과 파일이면 '05_계정명 리스트_통계' 시트를, 그 밖의 xlsx/csv는 첫 시트에서
    계정명 열을 찾는다. 금액 열은 읽지 않는다.
    """
    if path.suffix.lower() == ".csv":
        raw = pd.read_csv(path, header=None, dtype=object)
    else:
        xl = pd.ExcelFile(path)
        raw = xl.parse(STATS_SHEET if STATS_SHEET in xl.sheet_names else xl.sheet_names[0], header=None, dtype=object)

    company, header_row = None, None
    for i in range(min(len(raw), 15)):
        cells = [str(v).strip() for v in raw.iloc[i].tolist() if pd.notna(v)]
        m = re.match(r"회사명\s*:\s*(.+)", cells[0]) if cells else None
        if m:
            company = m.group(1).strip()
        if any(c in NAME_HEADERS for c in cells):
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"계정명 열을 찾지 못했습니다 (머리글 후보: {', '.join(NAME_HEADERS)})")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row].tolist()]
    name_col = next(c for c in df.columns if c in NAME_HEADERS)
    code_col = next((c for c in df.columns if c in CODE_HEADERS), None)
    out = pd.DataFrame({"account_name": df[name_col]})
    out["account_code"] = df[code_col] if code_col else None
    for field, header in (("n_debit", "차변건수"), ("n_credit", "대변건수"), ("n_vouchers", "전표개수")):
        out[field] = pd.to_numeric(df[header], errors="coerce") if header in df.columns else None
    out = out[out["account_name"].notna()].copy()
    out["account_name"] = out["account_name"].astype(str).str.strip()
    # "[10301]보통예금"처럼 코드가 붙은 계정명은 분리한다
    prefixed = out["account_name"].str.extract(r"^\[(\w+)\]\s*(.+)$")
    has_prefix = prefixed[1].notna()
    out.loc[has_prefix, "account_code"] = prefixed.loc[has_prefix, 0]
    out.loc[has_prefix, "account_name"] = prefixed.loc[has_prefix, 1]
    out = out[out["account_name"] != ""].drop_duplicates("account_name").reset_index(drop=True)
    return company, out


def _schema(categories: list[str]) -> dict:
    return {
        "type": "object",
        "properties": {"results": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "no": {"type": "integer"},
                "category": {"anyOf": [{"type": "string", "enum": categories}, {"type": "null"}]},
                "confidence": {"type": "string", "enum": list(LLM_CONFIDENCE)},
            },
            "required": ["no", "category", "confidence"],
            "additionalProperties": False,
        }}},
        "required": ["results"],
        "additionalProperties": False,
    }


def llm_classify(names: list[str], tax: Taxonomy, llm: LLM, log) -> dict[str, tuple[str | None, float]]:
    """계정명만 보내 의미 기반으로 분류한다. {계정명: (분류 또는 None, 신뢰도)}"""
    system = MAPPING_SYSTEM.format(categories=tax.described)
    schema = _schema(tax.names)
    out: dict[str, tuple[str | None, float]] = {}
    for start in range(0, len(names), BATCH):
        batch = names[start:start + BATCH]
        listing = "\n".join(f"{i}. {n}" for i, n in enumerate(batch, start=1))
        try:
            data = llm.call_json("mapping", llm.cfg["mapping_model"], system,
                                 [{"type": "text", "text": listing}], schema)
        except (LLMError, json.JSONDecodeError) as e:
            log.error("계정 의미 매칭 실패 (%d~%d번): %s", start + 1, start + len(batch), e)
            continue
        for item in data["results"]:
            if 1 <= item["no"] <= len(batch):
                out[batch[item["no"] - 1]] = (item["category"], LLM_CONFIDENCE[item["confidence"]])
    return out


def _classify(names: list[str], tax: Taxonomy, llm: LLM | None, threshold: float, log) -> dict[str, dict]:
    """이름 목록을 분류한다. {이름: {category, method, confidence, keyword}}"""
    result, ask = {}, []
    for name in names:
        hit = tax.match(name)
        if hit:
            result[name] = {"category": hit[0], "method": "사전", "confidence": hit[1], "keyword": hit[2]}
        else:
            result[name] = {"category": None, "method": None, "confidence": None, "keyword": None}
        # 글자가 하나도 없는 이름("0" 등)은 보내도 의미가 없다
        if (hit is None or hit[1] < threshold) and re.search(r"[A-Za-z가-힣]", name):
            ask.append(name)
    if llm and ask:
        for name, (category, confidence) in llm_classify(ask, tax, llm, log).items():
            if category:  # LLM이 null이면 사전 결과(있다면)를 그대로 둔다
                result[name] = {"category": category, "method": "LLM", "confidence": confidence,
                                "keyword": result[name]["keyword"]}
    return result


def map_findings(cfg: dict, store: Store, tax: Taxonomy, llm: LLM | None, log) -> None:
    """지적사항의 관련 계정·쟁점분야를 표준 분류로 정규화해 finding_categories에 저장한다."""
    findings = store.query("SELECT finding_id, related_accounts, issue_area FROM findings")
    terms_of: dict[str, list[str]] = {}
    for f in findings:
        terms = [*json.loads(f["related_accounts"] or "[]"), *([f["issue_area"]] if f["issue_area"] else [])]
        terms_of[f["finding_id"]] = [t for t in dict.fromkeys(terms) if _norm(t) not in tax.ignore]
    all_terms = sorted({t for terms in terms_of.values() for t in terms})
    classified = _classify(all_terms, tax, llm, cfg["mapping"]["llm_threshold"], log)

    store.conn.execute("DELETE FROM finding_categories")
    for fid, terms in terms_of.items():
        best: dict[str, dict] = {}
        for term in terms:
            c = classified[term]
            if c["category"] and (c["category"] not in best or c["confidence"] > best[c["category"]]["confidence"]):
                best[c["category"]] = {**c, "term": term}
        for category, c in best.items():
            store.upsert("finding_categories", {
                "finding_id": fid, "category": category, "source_term": c["term"],
                "match_method": c["method"], "match_confidence": c["confidence"]})
    store.commit()
    unmatched = [t for t in all_terms if not classified[t]["category"]]
    log.info("지적사항 분류: %d건, 용어 %d개 중 미분류 %d개%s", len(findings), len(all_terms), len(unmatched),
             f" ({', '.join(unmatched[:10])})" if unmatched else "")


def run_map(cfg: dict, store: Store, log, accounts_path: Path, company: str | None = None,
            use_llm: bool = True) -> None:
    tax = Taxonomy(cfg["taxonomy"])
    try:
        found_company, accounts = load_account_list(accounts_path)
    except (OSError, ValueError, StopIteration) as e:
        log.error("계정리스트를 읽지 못했습니다: %s (%s)", accounts_path, e)
        store.log("map", "ERROR", accounts_path.name, f"계정리스트 읽기 실패: {e}")
        return
    company = company or found_company or accounts_path.stem
    llm = LLM(cfg, store, log) if use_llm else None

    map_findings(cfg, store, tax, llm, log)

    classified = _classify(accounts["account_name"].tolist(), tax, llm, cfg["mapping"]["llm_threshold"], log)
    store.conn.execute("DELETE FROM account_mappings WHERE company=?", (company,))
    for row in accounts.itertuples(index=False):
        c = classified[row.account_name]
        store.upsert("account_mappings", {
            "company": company, "account_name": row.account_name,
            "account_code": None if pd.isna(row.account_code) else str(row.account_code),
            "category": c["category"], "match_method": c["method"] or "매칭 없음",
            "match_confidence": c["confidence"], "matched_keyword": c["keyword"],
            "n_debit": None if pd.isna(row.n_debit) else int(row.n_debit),
            "n_credit": None if pd.isna(row.n_credit) else int(row.n_credit),
            "n_vouchers": None if pd.isna(row.n_vouchers) else int(row.n_vouchers),
            "source_file": accounts_path.name, "updated_at": now(),
        })
    store.commit()

    methods = pd.Series([c["method"] or "매칭 없음" for c in classified.values()]).value_counts().to_dict()
    msg = f"{company}: 계정 {len(accounts)}개 → " + ", ".join(f"{k} {v}" for k, v in methods.items())
    log.info(msg)
    store.log("map", "INFO", accounts_path.name, msg)
    if llm:
        log.info(llm.usage_summary())
    export_mapping(cfg, store, log, company)


def export_mapping(cfg: dict, store: Store, log, company: str) -> Path:
    """매칭 결과 점검용 엑셀 (사전 보정에 쓴다)."""
    accounts = pd.read_sql_query(
        "SELECT a.account_name AS 계정명, a.account_code AS 계정코드, a.category AS 표준분류, "
        "a.match_method AS 매칭방식, a.match_confidence AS 신뢰도, a.matched_keyword AS 근거키워드, "
        "(SELECT COUNT(*) FROM finding_categories fc JOIN findings f ON f.finding_id=fc.finding_id "
        "  WHERE fc.category=a.category AND f.review_status<>'제외') AS 관련지적사항수, "
        "a.n_vouchers AS 전표개수 FROM account_mappings a WHERE a.company=? ORDER BY a.category, a.account_name",
        store.conn, params=(company,))
    findings = pd.read_sql_query(
        "SELECT fc.finding_id, f.title AS 제목, fc.category AS 표준분류, fc.source_term AS 근거용어, "
        "fc.match_method AS 매칭방식, fc.match_confidence AS 신뢰도 "
        "FROM finding_categories fc JOIN findings f ON f.finding_id=fc.finding_id ORDER BY fc.finding_id",
        store.conn)
    safe = re.sub(r'[\\/:*?"<>|]', "_", company)
    path = cfg["paths"]["output"] / f"mapping_{safe}.xlsx"
    matched = accounts[accounts["표준분류"].notna()]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        matched.to_excel(writer, sheet_name="계정 매칭", index=False)
        accounts[accounts["표준분류"].isna()][["계정명", "계정코드", "전표개수"]].to_excel(
            writer, sheet_name="매칭 없음", index=False)
        findings.to_excel(writer, sheet_name="지적사항 분류", index=False)
        for ws in writer.sheets.values():
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 14 if col[0].column > 1 else 34
            ws.freeze_panes = "A2"
    log.info("매칭 결과 점검용 엑셀: %s", path)
    return path
