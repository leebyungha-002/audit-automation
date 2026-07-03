# 세션 로그

## 2026-06-26 (2차)

**완료 작업**:
- 런처앱 [리스 분析] 리스 스케줄 생성 및 [파일 분류] 감사조서 파일 분류 도구에 회사 선택 콤보박스 추가
- 이전 세션의 은행조회서 완전성(메뉴 22) 관련 작업도 완료 상태로 인계됨

**변경 파일**:
- `launcher.py`: `detect_lease_companies()` 추가, lease_schedule `"company": "lease"`, file_classifier `"company": "js"` + `"extra": "company_flag"` 설정, `_on_tool_selected()` / `_run()` 분기 처리
- `lease_analyzer/lease_schedule.py`: `main()`에 `company` positional 인수 추가 (선택 회사 파일만 처리)
- `file_classifier/main.py`: `--company` 인수 추가, `AuditClassifierApp(company=...)` 전달 시 대상 폴더 자동 설정

**미해결 이슈**: 없음

**다음 할 일**:
1. 필요 시 기능 테스트 (런처앱 실행 후 리스·파일분류 도구에서 회사 선택 동작 확인)
2. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-06-26 (3차)

**완료 작업**:
- 감사주석 검증 도구(note_verifier) 개발 완료
  - C안(감사조서 마스터 템플릿) 채택: 블록1(좌)에 정산표/DSD 자동채움, 블록2(우)와 비교
  - PyQt6 GUI: 파일선택·소스유형(정산표/DSD)·소스단위(천원/원) 선택 → 검증결과 xlsx 저장
  - 블록 경계 자동 감지(빈 열 2개 연속), bool/문자열 셀 덮어쓰기 방지
  - launcher.py에 [주석 검증] JS/journal 두 항목 추가
- sejoong 정산표 구조 확인: 원 단위 저장, 시트명 1~33 감사조서와 일치

**변경 파일**:
- `note_verifier/note_verifier.py` (신규)
- `launcher.py` (주석 검증 도구 2개 추가)

**미해결 이슈**:
- note_verifier 정산표 단위 자동 감지 미구현
  - sejoong 정산표: 원 단위 (현재 사용자가 수동 선택)
  - 다른 회사 정산표: 천원 단위
  - **해법**: 소스 파일 숫자 중간값 > 10,000,000 이면 원 단위로 자동 판정 (÷1,000 적용)
  - 구현 위치: `run_verify()` 시작 부분에 `_detect_unit()` 함수 추가

**다음 할 일**:
1. note_verifier — 정산표 단위 자동 감지 구현 (`_detect_unit()`)
2. note_verifier — sejoong으로 실제 테스트 (감사조서 + 정산표_sejoong_25년.xlsx)
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-07-01

**완료 작업**:
- note_verifier 전면 개선 (총 6개 커밋)
  1. openpyxl 피벗 캐시 `formula=None` 버그 패치
  2. 시트별 진행률 표시줄 + 단계별 상태 레이블 추가
  3. 파일 로딩 단계별 상태 메시지 표시
  4. openpyxl → xlwings 전환 (Excel COM 기반, 고속 로딩)
  5. 로직 재설계: 라벨 매칭 방식 → 정산표 표 전체(라벨+값) 위치 기준 복사
  6. 표 순서 매칭 방식 적용: 빈 행 기준으로 표 탐지 후 순서대로 1:1 매핑
- 핵심 로직 변경사항:
  - 시트명 필터: 순수 숫자(`4`, `15`)만 대상, `4-1`·텍스트명 제외
  - 블록 탐지: 연속 2개 빈 열로 왼쪽/오른쪽 블록 구분
  - 표 탐지: 오른쪽 블록(감사인 작성) 기준 빈 행으로 표 위치 탐지
  - 복사: 정산표 표N → 감사조서 왼쪽 블록 표N 위치에 행 오프셋 적용하여 복사
  - 비교: 왼쪽(정산표 값) vs 오른쪽(감사인 값) 열 위치 기준 비교 + 색상

**변경 파일**:
- `note_verifier/note_verifier.py`

**미해결 이슈**:
- 실제 파일 테스트 미완료 (저녁에 계속 진행 예정)
- 정산표 단위 자동 감지 미구현 (이전 세션 이슈)

**다음 할 일**:
1. 실제 감사조서 + 정산표로 note_verifier 테스트
2. 블록/표 탐지 결과 확인 및 필요 시 로직 보정
3. 정산표 단위 자동 감지 구현 (`_detect_unit()`)
4. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-07-02~03

**완료 작업**:
- note_verifier 핵심 버그 수정 (총 4개 커밋)
  1. 감사조서 오른쪽 블록 단위 자동감지 제거 (천원단위 표는 항상 천원)
  2. 표별 블록 경계 탐지: `_find_blocks_in_range(ws, row_start, row_end)` 적용
  3. 감사조서 표 위치를 천원단위 블록 기준으로 탐지 (왼쪽 빈 블록 오인식 수정)
     - 2단계 탐지: 상단 15행으로 prelim_b2s 먼저 파악 → 천원단위 열 범위로 aud_s 탐지
  4. 왼쪽 블록 비어있을 때 copy_cols 오류 수정 (b1e 대신 `b2s - 3` 사용)
     - 약속 양식: 빈열2개 + 색깔열1개 = 3열 고정 간격
- 감사조서 양식 확정:
  - `[정산표 복사표] [빈열2개] [색깔구분열(값없음)] [천원단위 표] [색깔구분열] [원단위 표]`
  - 표마다 왼쪽 블록 열 수가 달라도 녹색 패딩으로 오른쪽 블록 시작 열 통일
  - 원단위 표 아래 계산내역 60행 이상이어도 정확히 탐지
- 단위 변환: 자동감지 불필요, GUI에서 수동 선택 (sejoong=원단위 선택)

**변경 파일**:
- `note_verifier/note_verifier.py`

**미해결 이슈**:
- 실제 파일 테스트 미완료 (런처앱 실행 전 중단)

**다음 할 일**:
1. 런처앱 실행 후 실제 감사조서 + sejoong 정산표로 note_verifier 테스트
2. 블록/표 탐지 결과 확인 및 필요 시 로직 보정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---


**완료 작업**:
- note_verifier 로직 개선 (총 3개 커밋)
  1. 감사조서 오른쪽 블록 단위 자동감지 제거 (천원단위 표는 항상 천원 → 오인식 방지)
  2. 표별 블록 경계 탐지 적용: `_find_blocks` → `_find_blocks_in_range(ws, row_start, row_end)` — 시트 내 표마다 열 너비가 달라도 각 표 행 범위 안에서 정확히 b1e/b2s 탐지
  3. 감사조서 표 탐지 시 왼쪽 4열만 스캔 (`col_end=4`) — 원단위 표 아래 계산내역이 60행 이상 이어져도 오인식 없이 좌측 표 기준으로만 탐지
- 감사조서 양식 확정:
  - `[정산표 복사표] [빈열 2개] [녹색 구분열(값 없음)] [천원단위 감사조서표] [노란색 구분열] [원단위 감사조서표]`
  - 표마다 왼쪽 블록 열 수가 달라도 녹색 패딩으로 통일 가능
  - 시트명 순수숫자(`4`, `15`)만 처리, `4-1`·텍스트명 제외

**변경 파일**:
- `note_verifier/note_verifier.py`

**미해결 이슈**:
- 실제 파일 테스트 미완료
- 정산표 단위 자동 감지 미구현

**다음 할 일**:
1. 실제 감사조서 + 정산표로 note_verifier 테스트
2. 블록/표 탐지 결과 확인 및 필요 시 로직 보정
3. 정산표 단위 자동 감지 구현 (`_detect_unit()`)
4. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---
