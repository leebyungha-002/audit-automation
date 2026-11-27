"""첨부파일 다운로드: 금감원 주소만 허용, 기존 파일은 덮어쓰지 않음"""
import hashlib
import html
import json
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests

ALLOWED_HOSTS = {"www.fss.or.kr", "fss.or.kr"}
MAX_BYTES = 100 * 1024 * 1024


class DownloadError(Exception):
    pass


def safe_name(name: str) -> str:
    """경로 문자와 Windows 금지 문자를 제거하고 길이를 제한"""
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", html.unescape(name or "")).strip(" .")
    suffix = Path(name).suffix[:10]
    stem = name[: len(name) - len(Path(name).suffix)][:80]
    return (stem or "attachment") + suffix


def _get(url: str) -> requests.Response:
    for _ in range(4):  # 이동(redirect)할 때마다 주소를 다시 검사
        p = urlparse(url)
        if p.scheme != "https" or p.hostname not in ALLOWED_HOSTS:
            raise DownloadError("금감원 https 주소가 아니어서 중단했습니다.")
        res = requests.get(url, timeout=60, stream=True, allow_redirects=False)
        if res.status_code not in (301, 302, 303, 307, 308):
            res.raise_for_status()
            return res
        url = urljoin(url, res.headers.get("Location", ""))
    raise DownloadError("주소 이동이 너무 많습니다.")


def download(api: str, item: dict, root: str = "downloads") -> Path:
    url = html.unescape(item.get("atchfileUrl") or "").strip()
    if not url:
        raise DownloadError("첨부파일 주소가 없습니다.")
    try:
        res = _get(url)
        data = bytearray()
        for chunk in res.iter_content(65536):
            data.extend(chunk)
            if len(data) > MAX_BYTES:
                raise DownloadError("파일이 100MB를 넘어 중단했습니다.")
    except requests.RequestException as e:
        raise DownloadError(f"다운로드 실패: {type(e).__name__}") from None
    if not data:
        raise DownloadError("빈 파일입니다.")
    if bytes(data[:200]).lstrip().lower().startswith((b"<!doctype html", b"<html", b"<script")):
        raise DownloadError("파일 대신 웹페이지가 왔습니다. 원문에서 확인하세요.")

    digest = hashlib.sha256(data).hexdigest()
    folder = Path(root) / safe_name(api) / safe_name(str(item["contentId"]))
    folder.mkdir(parents=True, exist_ok=True)
    # 파일 내용이 다르면 해시가 달라져 별도 파일로 보존
    path = folder / f"{digest[:12]}_{safe_name(item.get('atchfileNm'))}"
    if path.exists() and hashlib.sha256(path.read_bytes()).hexdigest() == digest:
        return path
    path.write_bytes(data)
    meta = {"url": url, "sha256": digest, "size": len(data),
            "downloaded_at": datetime.now().astimezone().isoformat(timespec="seconds")}
    path.with_name(path.name + ".meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
