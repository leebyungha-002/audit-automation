#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
journal_analyzer 수동 메뉴 선택 실행기
main_analyzer.py의 --task 옵션을 대화형으로 호출한다.
"""

import sys
import os
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)

MENU = {
    2:  '거래처비교',
    3:  '벤포드분석',
    4:  '데이터개요',
    5:  '계정명리스트',
    6:  '사원별집계',
    7:  '일자차이분석',
    8:  '상대계정분석',
    9:  '키워드검색',
    10: '라운드넘버',
    11: '특수관계자분석',
    12: '자산부채교차',
    13: '매출비용교차',
    14: '심층분석',
    15: 'AI계정별분석',
    16: '헤더확인',
    17: '거래처분석',
    18: '벤포드이탈',
    19: '월별전계정분석',
    20: '잔액증감분석',
    21: '총계정원장',
    22: '은행조회서완전성',
    23: '계정별상세내역',
}


def print_menu():
    print('\n' + '=' * 45)
    print('  분개장 분석 — 메뉴 선택')
    print('=' * 45)
    for no, name in MENU.items():
        print(f'  {no:>2}. {name}')
    print('=' * 45)
    print('  입력 예)  3 8 21   또는   all (전체)')
    print('=' * 45)


def parse_input(raw: str) -> list[int]:
    raw = raw.strip().lower()
    if raw == 'all':
        return list(MENU.keys())
    nums = []
    for token in raw.split():
        try:
            n = int(token)
            if n in MENU:
                nums.append(n)
            else:
                print(f'  [경고] {n}번은 없는 메뉴입니다. 건너뜁니다.')
        except ValueError:
            print(f'  [경고] "{token}"은 숫자가 아닙니다. 건너뜁니다.')
    return nums


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(here, 'main_analyzer.py')

    # 회사 선택
    company = input('\n고객사 이름 입력: ').strip()
    if not company:
        print('[오류] 회사 이름이 비어 있습니다.')
        sys.exit(1)

    while True:
        print_menu()
        raw = input('\n실행할 번호: ').strip()
        if not raw:
            continue

        tasks = parse_input(raw)
        if not tasks:
            print('  [오류] 유효한 번호가 없습니다. 다시 입력하세요.')
            continue

        selected = ', '.join(f'{n}. {MENU[n]}' for n in tasks)
        print(f'\n  실행: {selected}')

        cmd = [sys.executable, main_script, company, '--task'] + [str(n) for n in tasks]
        print(f'  명령: {" ".join(cmd)}\n')

        subprocess.run(cmd, cwd=here)

        again = input('\n계속 다른 메뉴를 실행하시겠습니까? (y/n): ').strip().lower()
        if again != 'y':
            break

    print('\n종료합니다.')


if __name__ == '__main__':
    main()
