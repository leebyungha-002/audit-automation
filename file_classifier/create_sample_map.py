"""
category_map.xlsx 샘플 생성 스크립트
------------------------------------
처음 사용 시 한 번만 실행하여 category_map.xlsx 템플릿을 만든다.
이후에는 엑셀 파일을 직접 편집하여 계정과목/키워드를 추가한다.

실행: python create_sample_map.py
"""

import os
import pandas as pd

# 감사 실무에서 자주 쓰이는 샘플 매핑 데이터
SAMPLE_DATA = [
    {
        '계정과목': '현금및현금성자산',
        '키워드':   '잔액증명,통장,은행,cash,bank,잔고,예금',
    },
    {
        '계정과목': '매출채권',
        '키워드':   '채권,receivable,invoice,외상매출,AR',
    },
    {
        '계정과목': '재고자산',
        '키워드':   '재고,inventory,창고,실사,stock',
    },
    {
        '계정과목': '유형자산',
        '키워드':   '유형자산,설비,기계,건물,차량,취득,감가상각,PP&E',
    },
    {
        '계정과목': '매입채무',
        '키워드':   '채무,payable,매입,AP,외상매입',
    },
    {
        '계정과목': '차입금',
        '키워드':   '차입,대출,loan,borrowing,금융기관',
    },
    {
        '계정과목': '매출',
        '키워드':   '매출,revenue,sales,세금계산서,계산서',
    },
    {
        '계정과목': '급여',
        '키워드':   '급여,임금,salary,payroll,인건비,급료',
    },
    {
        '계정과목': '법인세',
        '키워드':   '법인세,세금,tax,income tax,납부',
    },
    {
        '계정과목': '리스',
        '키워드':   '리스,lease,임대,사용권,ROU',
    },
]

def main():
    base_dir    = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(base_dir, 'category_map.xlsx')

    df = pd.DataFrame(SAMPLE_DATA, columns=['계정과목', '키워드'])
    df.to_excel(output_path, index=False)

    print(f"샘플 category_map.xlsx 생성 완료: {output_path}")
    print(f"총 {len(df)}개 계정과목 등록됨.")
    print("이 파일을 열어 계정과목과 키워드를 자유롭게 수정하세요.")

if __name__ == '__main__':
    main()
