#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_redactor/redact_pdf.py

PDF 문서에서 회사명·거래처명·계좌번호 등 민감정보를 지우거나 다른 이름으로 대체해
새 PDF / Word(.docx) 파일로 저장한다. (외부 전송 없이 로컬에서만 처리)

Usage:
    python redact_pdf.py                       # input/ 의 모든 PDF → PDF + Word
    python redact_pdf.py --format pdf          # PDF만 생성 (pdf | docx | both)
    python redact_pdf.py --scan                # 파일 생성 없이 민감정보 후보만 조사 → logs/
    python redact_pdf.py --file 계약서          # 파일명에 '계약서'가 포함된 PDF만
    python redact_pdf.py --rules rules/redact_rules_삼동.xlsx

규칙 파일 (rules/redact_rules.xlsx, 없으면 템플릿 자동 생성):
    [대체목록]  원문 | 대체어 | 사용(Y/N) | 비고
                → 대체어가 빈칸이면 검은 박스로 가림
    [자동패턴]  패턴 | 사용(Y/N) | 대체문구 | 설명
                → 계좌번호·사업자번호·전화번호 등 정규식 자동 탐지

처리 방식:
    PyMuPDF redaction으로 원문 텍스트를 PDF 내부에서 실제 삭제(가리기만 하는 것이 아님)한 뒤
    같은 위치에 대체어를 써 넣는다. 처리 후 결과 파일을 다시 읽어 원문 잔존 여부를 검증하고,
    처리내역·검증결과는 logs/ 에 기록한다 (logs/ 에는 원문이 남으므로 외부 전달 금지).
"""

import os
import re
import sys
import glob
import logging
import argparse
from datetime import datetime

import pandas as pd
import pymupdf
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
LOG_DIR    = os.path.join(BASE_DIR, 'logs')
RULES_DIR  = os.path.join(BASE_DIR, 'rules')
DEFAULT_RULES = os.path.join(RULES_DIR, 'redact_rules.xlsx')

FONT_CANDIDATES = [
    r'C:\Windows\Fonts\malgun.ttf',
    r'C:\Windows\Fonts\gulim.ttc',
]
FONT_NAME = 'krfont'

# ── 자동 탐지 패턴 (순서 = 우선순위, 먼저 잡힌 구간은 뒤 패턴에서 제외) ─────────────
# (패턴명, 정규식, 기본 대체문구, 기본 사용여부, 설명)
AUTO_PATTERNS = [
    ('이메일',
     r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+', '(이메일)', 'Y', 'abc@company.co.kr'),
    ('주민·법인등록번호',
     r'(?<![\d-])\d{6}-\d{7}(?![\d-])', '(등록번호)', 'Y', '000000-0000000'),
    ('사업자등록번호',
     r'(?<![\d-])\d{3}-\d{2}-\d{5}(?![\d-])', '(사업자번호)', 'Y', '000-00-00000'),
    ('카드번호',
     r'(?<![\d-])\d{4}-\d{4}-\d{4}-\d{4}(?![\d-])', '(카드번호)', 'Y', '0000-0000-0000-0000'),
    ('전화번호',
     r'(?<![\d-])0\d{1,2}[-)]\s?\d{3,4}-\d{4}(?![\d-])', '(전화번호)', 'Y', '02-000-0000, 010-0000-0000'),
    ('계좌번호',
     r'(?<![\d-])\d{2,6}(?:-\d{1,8}){1,4}(?![\d-])', '(계좌번호)', 'Y',
     '하이픈(-)으로 구분된 숫자 10~16자리 (날짜·금액은 제외됨)'),
    ('회사명(법인표기)',
     r'(?:\(주\)|㈜|주식회사)\s?[가-힣A-Za-z][가-힣A-Za-z0-9&]{1,20}'
     r'|[가-힣A-Za-z][가-힣A-Za-z0-9&]{1,20}\s?(?:\(주\)|㈜)',
     '회사{n}', 'N',
     '(주)/㈜/주식회사가 붙은 이름을 회사A, 회사B…로 일괄 대체. 조사까지 잡힐 수 있어 기본 N — '
     '--scan 결과의 [회사명후보]를 보고 대체목록에 직접 넣는 것을 권장'),
]
ACCOUNT_DIGITS = (10, 16)   # 계좌번호로 인정할 숫자 자릿수 범위


# ════════════════════════════════════════════════════════════════════════════
# 규칙 파일
# ════════════════════════════════════════════════════════════════════════════
def create_rules_template(path):
    wb = Workbook()
    hdr_font = Font(bold=True, color='FFFFFF')
    hdr_fill = PatternFill('solid', fgColor='305496')

    ws = wb.active
    ws.title = '대체목록'
    ws.append(['원문', '대체어', '사용', '비고'])
    ws.append(['(주)가나다상사', 'A사', 'N', '예시 행 — 사용=Y로 바꾸거나 지우고 실제 내용 입력'])
    ws.append(['가나다상사', 'A사', 'N', '법인표기 없는 표기도 따로 등록 (긴 원문부터 먼저 처리됨)'])
    ws.append(['홍길동', '', 'N', '대체어 빈칸 → 검은 박스로 가림'])
    for col, w in zip('ABCD', (30, 20, 8, 60)):
        ws.column_dimensions[col].width = w
    ws['A1'].comment = Comment('PDF에서 찾을 문자열. 띄어쓰기까지 정확히 입력', 'pdf_redactor')
    ws['B1'].comment = Comment('바꿔 넣을 문자열. 빈칸이면 검은 박스', 'pdf_redactor')
    ws['C1'].comment = Comment('Y/N (빈칸=Y)', 'pdf_redactor')

    ws2 = wb.create_sheet('자동패턴')
    ws2.append(['패턴', '사용', '대체문구', '설명'])
    for name, _, repl, use, desc in AUTO_PATTERNS:
        ws2.append([name, use, repl, desc])
    for col, w in zip('ABCD', (20, 8, 16, 90)):
        ws2.column_dimensions[col].width = w
    ws2['A1'].comment = Comment('패턴명은 수정하지 마세요 (코드와 연결됨)', 'pdf_redactor')
    ws2['C1'].comment = Comment('빈칸이면 검은 박스. 회사명(법인표기)의 {n}은 A,B,C…로 치환', 'pdf_redactor')

    for sh in (ws, ws2):
        for c in sh[1]:
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(horizontal='center')
        sh.freeze_panes = 'A2'

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def _yn(v):
    return str(v).strip().upper() not in ('N', 'NO', '0', 'FALSE')


def _cell(v):
    return '' if v is None else str(v).strip()


def load_rules(path):
    wb = load_workbook(path, data_only=True)

    terms = []
    if '대체목록' in wb.sheetnames:
        for row in wb['대체목록'].iter_rows(min_row=2, values_only=True):
            src, repl, use = (list(row) + [None] * 3)[:3]
            src = _cell(src)
            if src and (use is None or _yn(use)):
                terms.append((src, _cell(repl)))
    # 긴 원문 먼저 → "(주)가나다상사"가 "가나다상사"보다 먼저 처리되도록
    terms.sort(key=lambda t: len(t[0]), reverse=True)

    settings = {name: (use, repl) for name, _, repl, use, _ in AUTO_PATTERNS}
    if '자동패턴' in wb.sheetnames:
        for row in wb['자동패턴'].iter_rows(min_row=2, values_only=True):
            name, use, repl = (list(row) + [None] * 3)[:3]
            name = _cell(name)
            if name in settings:
                settings[name] = ('Y' if (use is None or _yn(use)) else 'N', _cell(repl))
    patterns = [(name, re.compile(rx), settings[name][1])
                for name, rx, *_ in AUTO_PATTERNS if settings[name][0] == 'Y']
    return terms, patterns


# ════════════════════════════════════════════════════════════════════════════
# 탐지
# ════════════════════════════════════════════════════════════════════════════
def _valid_match(name, s):
    if name == '계좌번호':
        n = sum(ch.isdigit() for ch in s)
        return ACCOUNT_DIGITS[0] <= n <= ACCOUNT_DIGITS[1]
    return True


def find_pattern_hits(text, patterns):
    """페이지 텍스트에서 자동패턴 매칭. 먼저 잡힌 구간은 다음 패턴에서 제외. [(패턴명, 문자열, 대체문구)]"""
    claimed, hits = [], []
    for name, rx, repl in patterns:
        for m in rx.finditer(text):
            s, e = m.span()
            if any(s < ce and e > cs for cs, ce in claimed):
                continue
            if not _valid_match(name, m.group()):
                continue
            claimed.append((s, e))
            hits.append((name, m.group().strip(), repl))
    return hits


class CompanyNamer:
    """회사명(법인표기) 패턴: 같은 이름은 실행 전체에서 같은 대체어(회사A, 회사B…)를 받는다."""
    def __init__(self):
        self.map = {}

    def get(self, name, template):
        key = re.sub(r'\(주\)|㈜|주식회사|\s', '', name)
        if key not in self.map:
            self.map[key] = _alpha_label(len(self.map))
        return template.replace('{n}', self.map[key]) if template else ''


def _alpha_label(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _overlaps(rect, claimed, ratio=0.5):
    area = rect.get_area() or 1
    return any((rect & c).get_area() / area > ratio for c in claimed)


def collect_targets(page, terms, patterns, namer):
    """페이지에서 지울 영역 수집. [(rect, 대체어, 규칙, 원문)]"""
    text = page.get_text('text')
    targets, claimed = [], []

    def add(needle, repl, rule):
        for r in page.search_for(needle):
            if _overlaps(r, claimed):
                continue
            claimed.append(r)
            targets.append((r, repl, rule, needle))

    for src, repl in terms:
        if src.lower() in text.lower():
            add(src, repl, '대체목록')

    seen = set()
    for name, s, repl in find_pattern_hits(text, patterns):
        if (name, s) in seen:
            continue
        seen.add((name, s))
        if name == '회사명(법인표기)':
            repl = namer.get(s, repl)
        add(s, repl, name)
    return targets


def is_scanned(doc):
    chars = sum(len(p.get_text('text').strip()) for p in doc)
    return chars < 20 * max(len(doc), 1)


# ════════════════════════════════════════════════════════════════════════════
# 처리
# ════════════════════════════════════════════════════════════════════════════
def _font_path():
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def write_replacement(page, rect, text, font, font_path):
    unit = font.text_length(text, fontsize=1) or 1
    # search_for 영역 높이는 줄 높이(글자 크기의 약 1.3~1.4배)라서 0.72배 ≈ 원문 글자 크기
    fs = max(min(rect.height * 0.72, rect.width * 0.98 / unit), 3)
    baseline = pymupdf.Point(rect.x0, rect.y0 + (rect.height + fs * 0.75) / 2)
    kw = {'fontfile': font_path} if font_path else {}
    page.insert_text(baseline, text, fontname=FONT_NAME if font_path else 'korea', fontsize=fs, **kw)


def apply_terms_to_text(s, terms):
    for src, repl in terms:
        s = s.replace(src, repl or '○○')
    return s


def redact_pdf(src, dst, terms, patterns, namer, font, font_path):
    doc = pymupdf.open(src)
    log_rows = []
    for page in doc:
        targets = collect_targets(page, terms, patterns, namer)
        if not targets:
            continue
        for rect, repl, _, _ in targets:
            page.add_redact_annot(rect, fill=(1, 1, 1) if repl else (0, 0, 0))
        # 이미지·표 테두리는 그대로 두고 텍스트만 삭제
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                              graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
        for rect, repl, rule, needle in targets:
            if repl:
                write_replacement(page, rect, repl, font, font_path)
            log_rows.append({'파일': os.path.basename(src), '페이지': page.number + 1,
                             '규칙': rule, '원문': needle, '대체어': repl or '(검은박스)'})

    # 메타데이터·책갈피에 남은 원문 제거
    doc.set_metadata({})
    doc.del_xml_metadata()
    toc = doc.get_toc(simple=True)
    if toc:
        doc.set_toc([[lv, apply_terms_to_text(t, terms), pg] for lv, t, pg in toc])
    doc.save(dst, garbage=4, deflate=True, clean=True)
    doc.close()
    return log_rows


def to_docx(pdf_path, docx_path):
    # pdf2docx 내부의 `import fitz`(폐지 예정 API, 경고 출력)를 pymupdf로 연결
    sys.modules.setdefault('fitz', pymupdf)
    from pdf2docx import Converter
    logging.getLogger().setLevel(logging.ERROR)
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path)
    finally:
        cv.close()


# ════════════════════════════════════════════════════════════════════════════
# 검증 — 결과 파일에 원문이 남아있지 않은지 다시 읽어 확인
# ════════════════════════════════════════════════════════════════════════════
def _pdf_text(path):
    with pymupdf.open(path) as d:
        return '\n'.join(p.get_text('text') for p in d) + '\n' + str(d.metadata)


def _docx_text(path):
    import docx
    d = docx.Document(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            parts.extend(c.text for c in row.cells)
    for s in d.sections:
        parts.extend(p.text for p in s.header.paragraphs)
        parts.extend(p.text for p in s.footer.paragraphs)
    return '\n'.join(parts)


def verify(text, needles, patterns):
    """{원문: (정확일치 잔존, 공백무시 잔존)} + 자동패턴 잔존"""
    low, squeezed = text.lower(), re.sub(r'\s', '', text).lower()
    res = {}
    for n in needles:
        exact = low.count(n.lower())
        loose = squeezed.count(re.sub(r'\s', '', n).lower())
        if exact or loose:
            res[n] = (exact, loose)
    for name, s, _ in find_pattern_hits(text, [p for p in patterns if p[0] != '회사명(법인표기)']):
        res.setdefault(s, (1, 1))
    return res


# ════════════════════════════════════════════════════════════════════════════
# 실행
# ════════════════════════════════════════════════════════════════════════════
COMPANY_CANDIDATE_RE = re.compile(AUTO_PATTERNS[-1][1])


def scan(pdfs, terms, patterns):
    """파일 생성 없이 후보만 조사."""
    rows, cands = [], {}
    all_patterns = [(n, re.compile(rx), r) for n, rx, r, *_ in AUTO_PATTERNS if n != '회사명(법인표기)']
    for src in pdfs:
        with pymupdf.open(src) as doc:
            if is_scanned(doc):
                rows.append({'파일': os.path.basename(src), '페이지': '-', '구분': '경고',
                             '문자열': '텍스트 없음(스캔본) — OCR 필요'})
                continue
            for page in doc:
                text = page.get_text('text')
                for src_t, _ in terms:
                    c = text.lower().count(src_t.lower())
                    if c:
                        rows.append({'파일': os.path.basename(src), '페이지': page.number + 1,
                                     '구분': '대체목록', '문자열': src_t, '건수': c})
                for name, s, _ in find_pattern_hits(text, all_patterns):
                    rows.append({'파일': os.path.basename(src), '페이지': page.number + 1,
                                 '구분': name, '문자열': s, '건수': 1})
                for m in COMPANY_CANDIDATE_RE.finditer(text):
                    cands[m.group().strip()] = cands.get(m.group().strip(), 0) + 1
    return pd.DataFrame(rows), pd.DataFrame(
        sorted(cands.items(), key=lambda x: -x[1]), columns=['회사명후보', '건수'])


def main():
    ap = argparse.ArgumentParser(description='PDF 민감정보 삭제·대체')
    ap.add_argument('--format', choices=['pdf', 'docx', 'both'], default='both')
    ap.add_argument('--rules', default=DEFAULT_RULES)
    ap.add_argument('--file', default='', help='파일명 포함 문자열 필터')
    ap.add_argument('--scan', action='store_true', help='파일 생성 없이 민감정보 후보만 조사')
    args = ap.parse_args()

    rules_path = args.rules if os.path.isabs(args.rules) else os.path.join(BASE_DIR, args.rules)
    for d in (INPUT_DIR, OUTPUT_DIR, LOG_DIR):
        os.makedirs(d, exist_ok=True)

    if not os.path.exists(rules_path):
        create_rules_template(rules_path)
        print(f'[안내] 규칙 파일이 없어 템플릿을 만들었습니다:\n  {rules_path}')
        print('  [대체목록] 시트에 회사명·거래처명 등을 입력한 뒤 다시 실행하세요.')
        print('  (먼저 --scan 으로 실행하면 문서 속 회사명 후보를 logs/ 에서 확인할 수 있습니다)')
        return

    terms, patterns = load_rules(rules_path)
    pdfs = sorted(p for p in glob.glob(os.path.join(INPUT_DIR, '*.pdf'))
                  if args.file in os.path.basename(p))
    if not pdfs:
        print(f'[안내] {INPUT_DIR} 에 처리할 PDF가 없습니다.')
        return

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    print(f'규칙: 대체목록 {len(terms)}건, 자동패턴 {", ".join(p[0] for p in patterns) or "없음"}')
    print(f'대상: PDF {len(pdfs)}개\n')

    if args.scan:
        df, cand = scan(pdfs, terms, patterns)
        out = os.path.join(LOG_DIR, f'scan_{ts}.xlsx')
        with pd.ExcelWriter(out) as xw:
            df.to_excel(xw, sheet_name='탐지내역', index=False)
            cand.to_excel(xw, sheet_name='회사명후보', index=False)
        print(f'탐지 {len(df)}건, 회사명후보 {len(cand)}개 → {out}')
        return

    font_path = _font_path()
    font = pymupdf.Font(fontfile=font_path) if font_path else pymupdf.Font('korea')
    namer = CompanyNamer()
    all_logs, checks, fails = [], [], 0

    for src in pdfs:
        name = os.path.basename(src)
        with pymupdf.open(src) as d:
            scanned = is_scanned(d)
        if scanned:
            print(f'[건너뜀] {name}: 텍스트가 없는 스캔본 PDF — OCR 후 다시 시도하세요')
            checks.append({'파일': name, '결과': '건너뜀(스캔본)'})
            continue

        stem = apply_terms_to_text(os.path.splitext(name)[0], terms)   # 파일명의 회사명도 대체
        out_pdf = os.path.join(OUTPUT_DIR, f'{stem}_비식별.pdf')
        logs = redact_pdf(src, out_pdf, terms, patterns, namer, font, font_path)
        all_logs.extend(logs)

        needles = {t for t, _ in terms} | {r['원문'] for r in logs}
        outputs = [('PDF', out_pdf, _pdf_text)]
        if args.format in ('docx', 'both'):
            out_docx = os.path.join(OUTPUT_DIR, f'{stem}_비식별.docx')
            try:
                to_docx(out_pdf, out_docx)
                outputs.append(('Word', out_docx, _docx_text))
            except Exception as e:
                print(f'  [오류] Word 변환 실패: {e}')
                checks.append({'파일': name, '형식': 'Word', '결과': f'변환실패: {e}'})

        for kind, path, reader in outputs:
            left = verify(reader(path), needles, patterns)
            ok = not left
            fails += 0 if ok else 1
            checks.append({'파일': name, '형식': kind, '결과파일': os.path.basename(path),
                           '결과': '통과' if ok else '잔존',
                           '잔존내역': '; '.join(f'{k}({a}/{b})' for k, (a, b) in left.items())})

        if args.format == 'docx':
            os.remove(out_pdf)
            outputs = outputs[1:]
        results = [c.get('결과', '') for c in checks if c['파일'] == name]
        status = ('⚠ 잔존 확인 필요' if '잔존' in results
                  else '⚠ 일부 변환실패' if any(r.startswith('변환실패') for r in results) else '통과')
        print(f'[완료] {name}: {len(logs)}곳 처리 → {", ".join(os.path.basename(o[1]) for o in outputs)}  [검증 {status}]')

    report = os.path.join(LOG_DIR, f'redact_report_{ts}.xlsx')
    with pd.ExcelWriter(report) as xw:
        pd.DataFrame(checks).to_excel(xw, sheet_name='검증', index=False)
        pd.DataFrame(all_logs).to_excel(xw, sheet_name='처리내역', index=False)
        if namer.map:
            pd.DataFrame([(k, f'회사{v}') for k, v in namer.map.items()],
                         columns=['원문', '대체어']).to_excel(xw, sheet_name='회사명대체표', index=False)
    print(f'\n처리내역·검증결과: {report}')
    print('  ※ logs/ 파일에는 원문이 들어있으니 외부로 전달하지 마세요. 전달용은 output/ 파일만 사용.')
    if fails:
        print(f'  ⚠ 원문이 남아있는 결과 {fails}건 — [검증] 시트의 잔존내역 확인 '
              '(괄호: 정확일치/공백무시 건수, 줄바꿈으로 끊긴 단어는 공백무시에서만 잡힘)')


if __name__ == '__main__':
    main()
