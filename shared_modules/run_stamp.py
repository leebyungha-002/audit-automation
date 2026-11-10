# -*- coding: utf-8 -*-
"""실행 PC·시각 기록 헬퍼 (미니PC/노트북 번갈아 실행 시 결과 파일 구분용).

- pc_label()       : '노트북(DESKTOP-V1M97UH)' — 별칭은 shared_modules/pc_alias.json
- stamp_workbook() : 엑셀 문서속성(만든 이/수정한 이/설명)에 실행정보 기록
- read_stamp()     : 엑셀 파일의 문서속성 설명(실행정보) 읽기
- append_history() : 텍스트 이력 파일에 한 블록 추가
"""
import os
import json
import socket
from datetime import datetime

_ALIAS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pc_alias.json')


def computer_name():
    return os.environ.get('COMPUTERNAME') or socket.gethostname()


def pc_label():
    name = computer_name()
    try:
        with open(_ALIAS_PATH, encoding='utf-8') as f:
            alias = json.load(f).get(name)
    except (OSError, ValueError):
        alias = None
    return f'{alias}({name})' if alias else name


def now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def mtime_str(path):
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')
    except OSError:
        return '?'


def stamp_workbook(wb, title):
    """openpyxl Workbook 문서속성에 실행정보 기록. 반환: 기록한 설명 문자열."""
    text = f'{title} | {now_str()} | {pc_label()}'
    props = wb.properties
    props.lastModifiedBy = pc_label()
    props.description = text
    props.modified = datetime.now()
    return text


def read_stamp(path):
    """xlsx 문서속성 설명(stamp_workbook이 남긴 실행정보)을 읽는다. 없으면 ''."""
    try:
        import zipfile
        import re
        with zipfile.ZipFile(path) as z:
            xml = z.read('docProps/core.xml').decode('utf-8', 'ignore')
        m = re.search(r'<dc:description[^>]*>(.*?)</dc:description>', xml, re.S)
        import html
        return html.unescape(m.group(1)) if m else ''
    except Exception:
        return ''


def append_history(log_path, lines):
    """이력 파일 끝에 블록 추가 (실패해도 본 작업은 계속)."""
    try:
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n\n')
        return True
    except OSError as e:
        print(f'  [경고] 실행이력 기록 실패: {e}')
        return False
