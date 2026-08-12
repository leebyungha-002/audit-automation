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

## 2026-07-06

**완료 작업**:
- note_verifier 표 인식 개선 (2가지 버그 수정)
  1. 표 선두 단위·설명 행 제거: `_trim_table_start()` 함수 추가
     - 소스(정산표): min_cols=3 — `"(1) 보고기간..."`, `"<당기말>"` 등 skip → 헤더 행부터 시작
     - 감사조서: min_cols=2, 오른쪽 블록 열 기준 — `"(단위:천원)"` 단일열 행 skip
     - 결과: src_s/aud_s가 헤더 행 기준으로 매칭 → 1행 오프셋 오류 해결
  2. bool 값 복사 방지: 정산표 TRUE/FALSE 셀 → None으로 변환 (감사조서에 TRUE/FALSE 쓰이는 것 방지)

**변경 파일**:
- `note_verifier/note_verifier.py`: `_trim_table_start()` 추가, `run_verify()` trim 적용, bool 필터

**미해결 이슈**:
- 실제 파일로 결과 확인 미완료 (런처앱 실행 전 세션 종료)

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 결과 확인
2. 표 복사 오프셋·내용 정확성 확인 후 필요 시 추가 보정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-07-09

**완료 작업**:
- 이전 세션(2026-07-06 2차) 사용량 한도 중단분 세션 로그 업데이트
  - git 이력 기준 4개 추가 커밋 확인 및 기록

**변경 파일**:
- `docs/session_log.md`

**미해결 이슈**:
- 실제 파일 테스트 미완료

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 결과 확인
2. 블록/표 탐지 결과 확인 및 필요 시 로직 보정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-07-09 (2차)

**완료 작업**:
- note_verifier — 첫 숫자 행 기준 정렬로 row 오프셋 오류 수정
  - `_first_numeric_row()` 헬퍼 추가: 지정 범위에서 첫 숫자 셀 행 반환
  - 비교 루프 변경: `src_r = src_s + i`, `right_r = aud_s + i` → `src_r = src_num_s + i`, `right_r = aud_num_s + i`
  - `write_r = right_r` (감사조서 시트 동일 행에 왼쪽 블록 쓰기)
  - 진단 로그: `src_num={} aud_num={} h={}` 출력으로 정렬 결과 확인 가능
  - 근본 원인: aud_s=헤더 행, src_num_s=첫 데이터 행으로 오프셋 불일치 → 모든 비교 어긋남

**변경 파일**:
- `note_verifier/note_verifier.py`: `_first_numeric_row()` 추가, 비교 루프 수정

**미해결 이슈**:
- 실제 파일 테스트 미완료
- 시트 1,2,3: `감사표:[]` (스퍼리어스 필터로 제거됨) — 테스트 후 재확인
- 시트 8,19,27: 블록 구분 열 미발견 — 단일 블록 시트 가능성

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 소차이/큰차이 분포 재확인
2. 잔여 이슈(시트 1-3 감사표 빈 배열, 시트 8/19/27 블록 미발견) 로그 확인 후 추가 수정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-07-21

**완료 작업**:
- journal_analyzer 음수 금액 추출 누락 수정 (4개 버그 동시 수정)
  1. `_to_numeric_amount()`: `(100,000)` 형식 음수 → 0으로 변환되던 버그 수정 (정규식 추가)
  2. 분석 함수 8곳의 `> 0` 필터 → `!= 0` 교체 (음수 행 제거 방지)
     - 사원별집계 / 상대계정 / 라운드넘버 / 심층분석 / 거래처분석 / 자산부채교차 / 매출비용교차 / 총계정원장 / 계정별상세내역
     - 벤포드 분석은 수학적으로 양수만 대상이므로 유지
  3. `라운드넘버`: 음수 금액도 `abs(x) % u` 로 정확히 탐지
- data_injector 매핑 리스트 컬럼 구조 문제 진단 및 안내
  - B열(src_kw): 소스 파일 키워드(`분석결과_kyungnam`) 필요 — 시트명 입력 시 파일 탐색 실패
  - D열(src_range): `B4:16` 형식 오류 → `B4:B16` (끝 셀도 열+행 형식 필요)
  - G열(start_cell) 비어있으면 행 전체 skip
- 총계정원장 — 데이터 없는 월 0으로 채우기 (단일 연도)
  - 기존: `groupby('YM')` 후 데이터 있는 월만 표시
  - 수정: 해당 연도 12개월 스켈레톤 생성 후 left join + fillna(0)
  - 다중 연도: 이미 `Month: range(1,13)` 스켈레톤 적용 중 → 변경 없음
- 총계정원장 계정 오매칭 수정 (핵심 버그)
  - 원인: `startswith(norm_user)` 로직으로 "급여 급료(제조)" 검색 시 "급여 급료(제조)의령"까지 포함
  - 증상: kyungnam 1월 급료제조 차변합계 364,601,641원 → 앱 413,111,245원 (차이 48,509,604 = 의령 계정분)
  - 수정: 자체 startswith 로직 → `_account_match_flexible()`(정확 일치 우선) 로 교체

**변경 파일**:
- `journal_analyzer/main_analyzer.py`

**미해결 이슈**: 없음

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 소차이/큰차이 분포 재확인
2. 잔여 이슈(시트 1-3 감사표 빈 배열, 시트 8/19/27 블록 미발견) 로그 확인 후 추가 수정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동 계속

---

## 2026-07-27

**완료 작업**:
- graphy JS 자동화 다수 버그 수정
  1. **임대료수입 다운로드 타임아웃**: `waitForEvent('download')` 예외를 try/catch로 처리, 경고 출력 후 skip
  2. **EBUSY 파일 잠금**: 그룹 파일 저장 시 temp→rename 15회 재시도 방식으로 OneDrive 잠금 대응
  3. **전기비교 기준월 미선택**: task_list 기준월 컬럼(6월) 읽어 `clickRadioByLabel`로 자동 클릭
  4. **벤포드 차트 소실**: `MASTER_MERGE_MENUS`에서 벤포드법칙분석 제외 → 계정별 개별 파일 저장
  5. **이중거래처분석 타임아웃**: `networkidle` 대기 + `handleDownloadAndSave` timeout 120초로 증가
  6. **총계정원장 이전 계정 재추출**: 계정 입력 후 `[role="option"]` count 확인, 0이면 skip
  7. **런처 로그 저장**: "로그 저장" 버튼 추가 (`docs/logs/launcher_log_YYYYMMDD_HHMMSS.txt`)
  8. **lease_filter 컬럼 공백**: `re.sub(r'\s+', '', ...)` 정규화 + `적요` astype(str) 수정
  9. **lease_filter 첫 계정만 분석**: 회사 모드에서 키워드 필터 비활성화 → 전 시트 처리
  10. **리스완전성 저장 경로**: `--output` 인수 추가 → `graphy/results/리스완전성_graphy.xlsx` 저장
  11. **런처 시트 필터**: "시트 필터" 입력 필드 추가 → `--sheet` 옵션으로 일부 메뉴만 실행
  12. **벤포드 금액기준열 코드로 됨(핵심)**:
      - 원인: UI 옵션이 "차 변"(공백 포함)인데 regex가 "차변"만 체크 → select 탐색 실패
      - 수정: `/차\s*변|대\s*변/.test(o)` 로 공백 무시 매칭 + 실제 레이블로 selectOption
      - 보조 전략 3: `clickRadioByLabel` fallback 추가, 대기시간 800ms → 1500ms

**변경 파일**:
- `shared_modules/auditRunner.js`
- `lease_analyzer/lease_filter.py`
- `launcher.py`

**미해결 이슈**: 없음

**다음 할 일**:
1. graphy 자동화 전체 재실행 (Excel 파일 닫고 실행) — 벤포드 기준열 + 전기비교 6월 기준 확인
2. sejoong / kyungnam mapping_list 작성 → data_injector 연동
3. 주석 검증(sejoong) 실행 → 소차이/큰차이 분포 재확인

---

## 2026-07-11

**완료 작업**:
- 세션 로그(2026-07-09 2차) 이후 추가 커밋 2개 확인:
  1. CLAUDE.md 오타(111---) 수정, 런처실행.bat 삭제
  2. note_verifier — NumberFormat 끝 쉼표로 원/천원 단위 자동 감지 구현 (57aec66)
     - `_detect_unit()`: 값 열 첫 금액 셀의 NumberFormat 끝 쉼표 확인 → 없으면 중간값 기반 폴백
     - `run_verify()` 시작부: 소스 단위 자동 감지 후 GUI 선택값 덮어씀

**변경 파일**:
- `docs/session_log.md`

**미해결 이슈**:
- 실제 파일 테스트 미완료 (런처앱 실행 필요)
- 시트 1,2,3: `감사표:[]` — 스퍼리어스 필터 과잉 제거 가능성
- 시트 8,19,27: 블록 구분 열 미발견 — 단일 블록 시트이거나 열 구조 다른 경우

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 소차이/큰차이 분포 재확인
2. 잔여 이슈(시트 1-3 감사표 빈 배열, 시트 8/19/27 블록 미발견) 로그 확인 후 추가 수정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-08-04

**완료 작업**:
- journal_analyzer 심층분석 파라미터 인식 버그 수정
  - 원인: 분析목록 분析명 "심층분析 (계정별 Top)" vs 시트명 "심층분析 (계정별 Top) " (trailing space 불일치)
  - 수정: `load_analysis_params()`에 `stripped_map` 추가 -> 앞뒤 공백 normalize 후 매칭
- journal_analyzer 은행조회서 완전성 파라미터 시트 추가
  - sejoong, kyungnam 두 task_list에 "은행조회서 완전성" 시트 추가 (기본 9개 계정)
  - 이후 task_list xlsx에서 직접 계정 추가/삭제 가능
- 런처앱 journal_analyzer 분析번호 선택 필터 추가
  - 분개장분析 도구 선택 시 "분析번호 선택" 입력란 표시
  - 번호 입력(예: `14` 또는 `3 8 14 22`) -> `--task` 인수로 전달 -> 지정 번호만 실행
  - 비워두면 task_list Y 항목 전체 실행 (기존 동작 유지)

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `load_analysis_params()` 시트명 공백 처리
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`: 은행조회서 완전성 시트 추가
- `journal_analyzer/sejoong/task_list_sejoong.xlsx`: 은행조회서 완전성 시트 추가
- `launcher.py`: 分析번호 필터 입력란 추가 (`_task_row`, `--task` 전달)

## 2026-08-04 (2차)

**완료 작업**:
- journal_analyzer 은행조회서 완전성 `str.contains` 경고 수정
  - 계정과목명에 괄호 `()`가 포함될 경우 pandas가 regex 그룹으로 오해 → `regex=False` 추가
- journal_analyzer 등록일자 파싱 `UserWarning` 수정
  - `pd.to_datetime(ser_r, errors='coerce')` → `format='mixed'` 추가 (pandas 2.0+)
  - 문자열 날짜가 다양한 형식으로 혼재해도 경고 없이 파싱
- graphy 원장 파일 업데이트 및 결과 파일 정리 (204개 파일 커밋)
- interest_analyzer/interest_expense_analysis.py 업데이트 커밋
- task_list_sejoong.xlsx, task_list_kyungnam.xlsx 커밋

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `regex=False`, `format='mixed'` 추가
- `graphy/raw_data/current/당기_graphy_계정별원장_26년2Q.xlsx`
- `interest_analyzer/interest_expense_analysis.py`
- `journal_analyzer/sejoong/task_list_sejoong.xlsx`
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`

**미해결 이슈**: 없음

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 소차이/큰차이 분포 재확인
2. 잔여 이슈(시트 1-3 감사표 빈 배열, 시트 8/19/27 블록 미발견) 로그 확인 후 추가 수정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-08-05

**완료 작업**:
- journal_analyzer 메뉴 25번 신규 추가: 손익월별분析
  - 손익 계정별 전기/당기 월별 비교 (1~12월 행, 증감금액·증감률%)
  - 거래처별 비교는 2번(거래처비교)과 중복 → 25번에서 제거, 월별만 출력
  - 레지스트리명: 손익항목分析 → 손익월별分析 변경
  - sejoong·kyungnam task_list 양쪽에 분析목록 25번 행 + 손익월별分析 파라미터 시트 추가
- 17번 거래처분析 파라미터 시트 추가 (sejoong·kyungnam 양쪽)
  - 컬럼: 작업명 / 계정과목 / 거래처명 / 금액열 / 실행여부
- sejoong 분析목록에 기준월 컬럼 추가 (kyungnam과 동일 구조)
  - 기준월은 분析함수 호출 전 공통 필터로 적용됨 (25번 포함 모든 메뉴에 적용)
- audit-automation Python 코드 → Pandas_Accounting_Tool 폴더로 복사 (15개 파일)
  - journal_analyzer, interest_analyzer, lease_analyzer, note_verifier, file_classifier, report, launcher.py

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: analyze_pl_comparison() 추가, 거래처 시트 제거
- `journal_analyzer/sejoong/task_list_sejoong.xlsx`: 25번·17번 시트 추가, 기준월 컬럼 추가
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`: 25번·17번 시트 추가

**미해결 이슈**: 없음

**다음 할 일**:
1. 런처앱 실행 → 주석 검증(sejoong) 실행 → 소차이/큰차이 분포 재확인
2. 잔여 이슈(시트 1-3 감사표 빈 배열, 시트 8/19/27 블록 미발견) 로그 확인 후 추가 수정
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동

---

## 2026-08-11 15:10

**완료 작업**: 리스 상각 스케줄에 계약별 상각 인식월(당월/익월) 선택 컬럼 추가
- 계약별로 상각을 리스개시월 당월부터 인식하거나 익월부터 인식하는 관행이 혼재되어 일괄 코딩이 어려운 문제 → input_data에 '상각개시(당월/익월)' 컬럼(드롭다운) 신설
- build_schedule()에서 이 값에 따라 월별 라벨(연월)만 1개월 오프셋, 총 기간 수·이자계산 로직은 불변
- 계약별 시트 상단 계약정보 영역에도 '상각개시' 값 표시
- 미입력 시 기존과 동일하게 '당월' 처리(하위호환)
- kyungnam·graphy input_data 파일에 새 컬럼(드롭다운 "당월,익월") 추가, 기존 계약은 전부 '당월' 기본값으로 채움

**변경 파일**:
- `lease_analyzer/lease_schedule.py`
- `lease_analyzer/input_data/lease_kyungnam_information_fy26.xlsx`
- `lease_analyzer/input_data/lease_graphy_information_fy25.xlsx`
- `lease_analyzer/output/lease_schedule_graphy_2025.xlsx` (재생성 검증)

**미해결 이슈**: kyungnam 결과 파일(`output/lease_schedule_kyungnam_2026.xlsx`)이 사용자 PC에서 엑셀로 열려있어 재생성 스크립트 실행이 PermissionError로 실패. 사용자가 파일을 닫은 후 재실행 필요.

**다음 할 일**:
1. kyungnam 엑셀 파일 닫은 뒤 `python lease_schedule.py --file lease_kyungnam_information_fy26.xlsx --fiscal-month 6` 재실행하여 결과 갱신
2. 실제 계약 중 익월 상각 대상이 있는지 확인 후 input_data에서 해당 계약만 '익월'로 표시
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동 (이전 세션 이월 항목)

---

## 2026-08-11 (2차)

**완료 작업**: journal_analyzer 메뉴15(AI계정별분석) 뒷단에 실제 Gemini API 호출 단계를 붙여 mapping_list로 감사조서에 자동 반영되는 구조 구현
- 메뉴15(`analyze_ai_preparation`)의 계정별 필터링/월별집계/샘플링/마스킹 로직을 `_prepare_ai_material()`로 공용화
- 신규 메뉴26 `AI계정별분석_실행`(`analyze_ai_review`) 추가: 계정과목별로 Gemini(`google-genai`, 구 `google-generativeai`는 지원종료 확인되어 신규 SDK로 채택)에 공통 JSON 스키마(위험평가/주요특이사항/결론/결론근거/추가확인사항)로 구조화 응답 요청 → 계정당 1행 표(`AI검토결과`) 생성
- `data_injector.py`에 `AI_INJECT` remarks 키워드 추가(`inject_ai_result`): 위험평가=높음 또는 결론=추가확인필요 행 노란색 강조 후 감사조서에 주입 (기존 `ANALYSIS_INJECT`와 동일 패턴)
- `.env`/`.env.example`에 `GEMINI_API_KEY`/`GEMINI_MODEL` 항목 추가, `python-dotenv`로 로드
- import/스키마 검증 등 오프라인 스모크 테스트 통과 (실제 API 키로 살아있는 호출은 미검증 — 키 미보유)

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `_prepare_ai_material`, `analyze_ai_review`, `_get_gemini_client_and_config`, `_build_ai_review_prompt`, `_AI_REVIEW_SCHEMA` 추가, `ANALYSIS_REGISTRY[26]` 등록, `load_dotenv()` 추가
- `report/data_injector.py`: `inject_ai_result()` 추가, remarks 분기에 `AI_INJECT` 연결
- `.env`, `.env.example`: `GEMINI_API_KEY`, `GEMINI_MODEL` 플레이스홀더 추가

**미해결 이슈**: 실사용 전 사용자가 직접 해야 할 것
1. `.env`의 `GEMINI_API_KEY`에 실제 키 입력
2. 대상 회사 `task_list_<회사>.xlsx`의 분석목록 시트에 26번 행 추가 + `AI계정별분석_실행_파라미터`(또는 분석목록에 적은 이름 그대로 + `_파라미터`) 시트에 계정과목 기입
3. 해당 회사 `mapping_list*.xlsx`에 신규 행 추가 — src_kw=`AI계정별분석_실행`, src_sheet=`AI검토결과`, remarks=`AI_INJECT`(강조 원할 시), tgt는 실제 감사조서 좌표

**다음 할 일**:
1. 사용자가 GEMINI_API_KEY 발급/입력 후 `python main_analyzer.py <회사> --task 26` 1회 시험 실행 (계정 1~2개만)
2. 결과 확인 후 실제 감사조서 좌표를 알려주면 mapping_list 행을 대신 채워줄 수 있음
3. sejoong / kyungnam mapping_list 작성 → data_injector 연동 (이전 세션부터 이월된 항목, 위 신규 AI_INJECT 행과 함께 정리 가능)

---

## 2026-08-11 (3차)

**완료 작업**: 실제 API 키로 메뉴26을 돌려보며 발견된 문제 2건 수정 — 둘 다 실제 Gemini 호출로 E2E 검증 완료
1. **모델 404 수정**: `gemini-2.5-flash`가 신규 발급 API 키에는 호출이 막혀있음(`models.list()`엔 나오지만 실제 generateContent는 404 "no longer available to new users") → 기본 모델을 `gemini-flash-lite-latest`(별칭, 향후 세대교체에도 자동 대응)로 교체
2. **AI검토결과 상세화**: "계정당 1행 요약만으론 어떤 전표를 확인해야 할지 모르겠다"는 피드백 반영 — JSON 스키마에 `확인필요전표`(전표번호+확인사유) 배열 추가, AI가 지목한 전표번호를 원본 샘플 데이터와 대조해 실제 금액·거래처·일자를 채운 `AI검토_확인전표` 시트를 별도 생성(확인필요 전표가 있을 때만 생성됨). AI가 응답에 적은 숫자는 신뢰하지 않고 항상 원본 대조.
   - 검증: 정상 케이스(이상거래 없음) → 확인전표 시트 미생성 / 대표이사 앞 1억원 이체 주입 케이스 → 정확히 해당 전표 1건만 잡아냄

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `_get_gemini_client_and_config` 기본 모델 변경, `_AI_REVIEW_SCHEMA`에 `확인필요전표` 배열 필드 추가, `_build_ai_review_prompt` 지시문 보강, `analyze_ai_review()` 반환 타입 **DataFrame → dict**로 변경(`AI검토결과` + 조건부 `AI검토_확인전표`)
- `.env.example`: `GEMINI_MODEL` 기본값 갱신 (`.env`는 gitignore 대상이라 로컬에서만 별도 수정됨)

**미해결 이슈**: mapping_list 작업은 내일 이어서 진행 예정 (사용자 요청)
- 아직 어느 회사부터 할지, 감사조서의 실제 셀 좌표가 무엇인지 확정되지 않음

**다음 할 일** (내일 시작점):
1. mapping_list 작업 대상 회사 확정 (sejoong / kyungnam 중, 혹은 다른 회사)
2. 해당 회사 `mapping_list*.xlsx`에 신규 행 2개 추가
   - `AI검토결과` → 감사조서 좌표, remarks=`AI_INJECT` (위험평가=높음/결론=추가확인필요 행 노란 강조)
   - `AI검토_확인전표` → 감사조서 좌표, remarks 비움 (표 그대로 주입, 확인필요 전표 있을 때만 시트 존재하므로 없으면 해당 행은 "소스 시트 없음" 처리될 수 있음 — 안내 필요)
3. 대상 회사 task_list 분석목록에 26번 행 + 파라미터 시트(계정과목) 아직 미등록 상태면 같이 세팅
4. `python report/data_injector.py <회사>` 실행해서 `_updated.xlsx`에 정상 반영되는지 확인
5. (이전부터 이월) sejoong / kyungnam mapping_list 전반 작성 — 위 AI_INJECT 행과 함께 한 번에 정리 가능

---

## 2026-08-11 (4차 - 세션 중단 복구)

**완료 작업**: 이전 세션이 중단된 지점을 진단·복구
- `interest_analyzer/interest_expense_analysis.py` 1행이 `11"""`로 깨져 있어 SyntaxError로 전혀 실행 불가한 상태 확인 (커밋되지 않은 상태, mtime 21:29 — 로그에 기록되지 않은 세션이 이 파일을 고치다 중단된 것으로 추정)
  - `git show HEAD:...`로 확인한 결과 원래 커밋(`6d064fa`)에도 `1"""`로 이미 깨져 있었음 → 이번 기회에 근본 오류까지 함께 수정
  - `"""`로 수정 후 `py_compile` 통과 확인, `python interest_expense_analysis.py dae_il` 실제 실행 → `이자비용분석결과.xlsx` 정상 생성 확인
  - ⚠️ 실행 결과 참고: dae_il 기대이자비용 164,940,000원 vs 장부상 이자비용 1,745,268,147원 (차이 +90.55%) — 코드 문제 아님, 감사인이 직접 확인 필요한 큰 차이
- session_log.md 마지막 항목(3차)의 미해결 이슈는 mapping_list 작성이었으나, 실제 중단 지점은 이 파일이었음 → mapping_list 작업은 아직 미착수 상태로 남아있음

**변경 파일**:
- `interest_analyzer/interest_expense_analysis.py`: 1행 `11"""` → `"""` 수정

**미해결 이슈**:
- 3차 세션의 mapping_list 작업(sejoong 신규 작성 + kyungnam AI_INJECT 행 추가)은 그대로 남아있음
- dae_il 이자비용 차이(+90.55%) 원인 미확인 — 감사인 검토 필요

**다음 할 일**:
1. mapping_list 작업 재개 여부/우선순위 사용자 확인 (interest_analyzer 건과 별개로 계속 이월 중)
2. dae_il 이자비용 차이 원인 확인 (등록 안 된 차입금 존재 여부 등)

---

## 2026-08-11 (5차)

**완료 작업**: kyungnam AI계정별분석_실행(26번) 세팅 + mapping_list 연동, 실제 실행으로 E2E 검증하며 핵심 버그 1건 발견·수정
- 사용자 결정사항: kyungnam 대상 계정=보통예금, 주입 위치=감사조서 신규 전용 시트 'AI검토' (AI검토결과 @A1, AI검토_확인전표 @A20)
- `task_list_kyungnam.xlsx`: 분석목록에 26번(AI계정별분석_실행, Y, 당기) 행 추가 + 'AI계정별분석_실행' 파라미터 시트(보통예금/Y) 신규 생성
- `Kyungnam_mapping_list_25년.xlsx`: AI검토결과/AI검토_확인전표 매핑 행 2개 추가 (둘 다 remarks=AI_INJECT)
- **핵심 버그 발견·수정**: `data_injector.py`의 `inject_analysis_result`/`inject_ai_result`가 분석결과 파일을 pandas로 읽을 때 1~2행(회사명/분석기간, `save_results()`가 항상 삽입)을 헤더로 오인식 → 실제 컬럼('위험평가','결론' 등)이 데이터로 밀려 강조 로직·주입 데이터가 전부 깨지는 상태였음 (ANALYSIS_INJECT 사용 매핑 행도 동일 결함 보유했을 것으로 추정되나 kyungnam엔 아직 해당 remarks 사용 행 없어 미발현 상태였음). `header=2` 추가로 수정, 실제 파일로 재현 후 수정 확인
- 검증 순서: kyungnam 전체 분석 실행(`main_analyzer.py kyungnam`, 125개 시트 생성, 26번의 Gemini 실호출 포함) → `data_injector.py kyungnam` 실행(46건 중 42건 성공) → `당기_Kyungnam_조서_26년_updated.xlsx`의 'AI검토' 시트를 직접 열어 헤더/데이터가 정확히 들어갔는지 확인 완료
- AI검토_확인전표는 이번 실행에서 이상거래가 없어 시트 자체가 생성되지 않아 "소스 시트 없음" 오류로 스킵됨 — 설계대로 동작(문제 아님)
- 이번 실행에서 함께 드러난 기존(무관) 오류 3건은 손대지 않음: 상세거래_구축물_차변/상세거래_시설장치_대변/상세거래_임대보증금(비유동)_대변 소스 시트 없음 — 해당 계정 당기 거래 없음으로 추정, 별도 확인 필요
- xlwings 재저장(XML 정합성 복구) 단계 실패 경고 있었으나 openpyxl 저장은 정상 완료되어 `_updated.xlsx` 자체는 사용 가능

**변경 파일**:
- `report/data_injector.py`: `inject_analysis_result`/`inject_ai_result` 내부 `pd.read_excel(...)` 에 `header=2` 추가
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`: 26번 등록 + 파라미터 시트 추가
- `journal_analyzer/kyungnam/감사조서/Kyungnam_mapping_list_25년.xlsx`: AI_INJECT 매핑 행 2개 추가

**미해결 이슈**:
- sejoong mapping_list는 아직 미작성 (파일 자체가 없음, 다음 세션에서 진행 예정)
- 이번에 드러난 kyungnam 기존 오류 3건(구축물/시설장치/임대보증금(비유동) 소스 시트 없음) 원인 미확인
- dae_il 이자비용 차이(+90.55%) 원인 미확인 (4차 세션 이월)

**다음 할 일**:
1. sejoong mapping_list 신규 작성 (감사조서 좌표 확인 필요 — 회사 형식이 kyungnam과 다를 수 있음) — 사용자 요청으로 보류 중
2. kyungnam 소스 시트 없음 오류 3건 원인 확인 (해당 계정 당기 거래 유무 확인)
3. dae_il 이자비용 차이 원인 확인

---

## 2026-08-11 (6차)

**완료 작업**: journal_analyzer 27번 메뉴 신규 추가 — 감가상각_평가손익분석 (Phase 1)
- 사용자 요청: 현금흐름표·주석 작성을 위해 감가상각비/외화환산손익/평가손익/대손상각비 등
  손익 계정의 상대계정(자산 관련 누계액 등) 금액을 8번(상대계정분석)과 같은 방식으로 찾고,
  이후 유형자산별 취득원가·감가상각누계액 롤포워드(기초+증가-감소=기말) 표까지 만들고 싶다는 요청
- 설계 논의 후 사용자 결정사항:
  - 기초잔액은 전기 계정별_거래처별명세 파일(회사별 data/previous/)이 있으면 그 파일 계정별 합산 사용,
    없으면(kyungnam처럼) 기초잔액은 표시하지 않음 — 분개장에서 억지로 역산하지 않기로 함
  - 구현은 1단계(상대계정 매칭)만 먼저 진행하기로 확정, 유형자산 롤포워드 표는 Phase 2로 보류
- `analyze_depreciation_valuation()` 추가 — 실제로는 8번 `analyze_counterpart()`를 그대로 재사용하는
  얇은 래퍼(동일 로직 요청이었으므로 중복 구현하지 않음), `ANALYSIS_REGISTRY[27]` 등록
- kyungnam·sejoong 양쪽 task_list에 27번 분석목록 행(실행여부=N, 검토 후 켜도록) +
  '감가상각_평가손익분석' 파라미터 시트(예시 계정 5개 시딩, 실제 계정과목명은 회사별 조정 필요) 추가
- kyungnam 실 데이터로 `--task 27` 실행 검증: 감가상각비 → 건물/기계장치/구축물/차량운반구/
  비품/시설장치/사용권자산/투자부동산_건물 감가상각누계액 등 정상 추출, 대손상각비 → 외상매출금/
  받을어음대손충당금 정상 추출. 단, 감가상각비는 전표번호 기준 그룹핑 특성상(8번과 동일 로직) 같은
  전표에 묶인 무관한 계정(보통예금, 재고자산 등)도 다수 함께 잡혀 노이즈가 큼 — 사용자가 결과 시트에서
  실제 감가상각누계액류만 걸러봐야 함 (기존 8번의 알려진 한계, 이번에 새로 만든 문제 아님)
  외화환산손익/공정가치평가손익/파생상품평가손익 3개 시딩 계정명은 kyungnam 실제 계정과목과 불일치해
  결과 시트 미생성 — 회사별 실제 계정명으로 교체 필요
- 테스트 후 kyungnam 27번 실행여부는 다시 N으로 원복(정식 실행 아니었으므로)

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `analyze_depreciation_valuation()` 추가, `ANALYSIS_REGISTRY[27]` 등록, 상단 파라미터 규격 주석 추가
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`: 27번 행 + 파라미터 시트 추가
- `journal_analyzer/sejoong/task_list_sejoong.xlsx`: 27번 행 + 파라미터 시트 추가

**미해결 이슈**:
- Phase 2 미구현: 유형자산별(건물/기계장치/구축물/차량운반구/비품/시설장치 등) 취득원가·감가상각누계액
  롤포워드 표(기초+당기증가-당기감소=기말) — 21번 총계정원장 활용 아이디어 사용자 제안, 다음 세션에서 설계 이어가기
- sejoong mapping_list 작성은 사용자 요청으로 계속 보류 중
- kyungnam 소스 시트 없음 오류 3건, dae_il 이자비용 차이 원인 확인 — 이월 지속

**다음 할 일**:
1. Phase 2 설계: 유형자산별 취득원가/감가상각누계액 롤포워드 표
   - 기초잔액: 전기 계정별_거래처별명세 있는 회사만 계정별 합산, 없으면 공란
   - 당기증가/당기감소: 21번 총계정원장 방식 활용 검토 (해당 유형자산 계정의 당기 차/대변)
   - 당기 감가상각비는 27번 Phase 1의 상대계정 매칭 결과를 자산 카테고리별로 재집계해서 연결
2. kyungnam/sejoong 27번 파라미터 시트의 실제 계정과목명 확인 후 정정 (외화환산손익 등 3개 계정명 불일치)
3. (이월) sejoong mapping_list, kyungnam 소스 시트 없음 오류 3건, dae_il 이자비용 차이

---

## 2026-08-11 (7차)

**완료 작업**: 27번 메뉴 Phase 2 — 유형자산별 취득원가/감가상각누계액 롤포워드 표 구현
- `_depreciation_rollforward()` 추가: task_list의 '감가상각_유형자산롤포워드' 시트(유형자산계정명/
  감가상각누계액계정명 쌍)를 읽어 계정별로 취득원가(기초+당기증가-당기감소=기말)와
  감가상각누계액(기초+당기감가상각비-당기감소=기말) 표 생성
  - 당기증가/당기감소: 해당 계정 자체의 당기 차/대변 합계 (직접 조회, 21번과 같은 방식)
  - 당기감가상각비: Phase 1의 '상대_감가상각*' 매칭 결과에서 해당 감가상각누계액계정명 금액을 가져옴
    (감가상각누계액 계정 자체의 대변 합계를 쓰지 않고 굳이 Phase1 매칭분을 쓴 이유: 상대계정 전표
    매칭으로 실제 감가상각비 발생분만 잡고 다른 대변 조정은 배제하기 위함 — 사용자 요청사항)
  - 기초잔액: 전기 계정별_거래처별명세 파일 있으면 계정 전체 합산, 없으면 공란(None)
- 리팩터링: 20번(잔액증감분석)에 있던 전기명세 탐색/로드 로직(파일탐색·시트매칭·잔액합산)을
  모듈 레벨 공용 헬퍼(`_find_prev_detail_file`/`_find_prev_sheet`/`_load_prev_balances`/
  `_prev_balance_total`)로 추출 — 20번 동작은 완전히 동일하게 유지하면서 27번과 공유
- kyungnam·sejoong task_list에 '감가상각_유형자산롤포워드' 파라미터 시트 추가
  - kyungnam: 실제 분개장 계정명(건물/구축물/기계장치/차량운반구/시설장치/비품/사용권자산/
    투자부동산_건물 + 상각 없는 토지/건설중인자산)으로 확인 후 시딩
  - sejoong: 예시 계정명(검증 필요)으로 시딩
- kyungnam 실 데이터로 `--task 27` 검증(테스트 후 실행여부 다시 N으로 원복):
  '유형자산_롤포워드' 시트 정상 생성, 기초잔액은 예상대로 공란(kyungnam에 전기명세 파일 없음)
  - **감사 관점에서 눈에 띄는 점**: 구축물/기계장치/차량운반구/시설장치/비품/사용권자산 6개 계정은
    '상각누계액_당기감가상각비'와 '상각누계액_당기감소' 금액이 정확히 일치함 (예: 기계장치 175,156,825원
    양쪽 동일). 실제 처분이 없었다면 이 회사 ERP가 상각누계액을 매월 대체/재기표하는 방식일 가능성 있음.
    건물·투자부동산_건물 2개 계정만 두 금액이 다름(실제 증감 반영 추정). 코드 버그로 보이지 않고 실제
    분개 데이터 특성으로 판단되나, 감사인이 직접 확인 필요 — 결과 그대로 보고, 임의로 보정하지 않음

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `_depreciation_rollforward()`, 전기명세 공용 헬퍼 4개 추가,
  `analyze_balance_movement()`가 공용 헬퍼 사용하도록 리팩터링, 상단 파라미터 규격 주석 갱신
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`: '감가상각_유형자산롤포워드' 시트 추가
- `journal_analyzer/sejoong/task_list_sejoong.xlsx`: '감가상각_유형자산롤포워드' 시트 추가

**미해결 이슈**:
- kyungnam 6개 계정의 당기감소=당기감가상각비 일치 현상 원인 미확인 — 감사인 판단 필요
- kyungnam/sejoong 27번 Phase 1 파라미터의 외화환산손익 등 3개 계정명 여전히 실제 계정과목과 불일치
- sejoong mapping_list 작성 계속 보류(사용자 요청), kyungnam 소스 시트 없음 오류 3건, dae_il 이자비용 차이 — 이월 지속

**다음 할 일**:
1. kyungnam 당기감소=당기감가상각비 일치 현상 사용자 확인 (ERP 재기표 방식인지 실제 처분 부재인지)
2. 27번 Phase 1 파라미터 계정명 실제 값으로 정정 (kyungnam/sejoong 공통)
3. (이월) sejoong mapping_list, kyungnam 소스 시트 없음 오류 3건, dae_il 이자비용 차이

---

## 2026-08-11 세션 종료 (내일 인계)

사용자 요청으로 오늘 세션 종료. 미해결 이슈 없음(코드 관점에서 막힌 것은 없고, 전부 사용자 확인/판단 대기 상태) — 새 세션은 아래 우선순위로 이어서 진행.

**내일 시작점 (우선순위순)**:
1. **kyungnam 27번 Phase 2 검증 결과 논의** — 구축물/기계장치/차량운반구/시설장치/비품/사용권자산 6개 계정에서 '당기감소'가 '당기감가상각비'와 정확히 일치하는 현상 (건물·투자부동산_건물만 다름). ERP가 감가상각누계액을 매월 재기표하는 방식인지, 아니면 코드 쪽에서 놓친 케이스가 있는지 사용자와 확인
2. **27번(감가상각_평가손익분석) 파라미터 계정명 정정** — kyungnam/sejoong 둘 다 외화환산손익/당기손익-공정가치측정금융자산평가손익/파생상품평가손익 3개는 예시로 시딩한 이름이라 실제 계정과목명과 불일치 (감가상각비·대손상각비는 kyungnam 기준 확인 완료). 확정되면 27번 분석목록 실행여부도 Y로 전환
3. **sejoong mapping_list 신규 작성** — 사용자가 "나중에"로 미룬 항목, 파일 자체가 아직 없음. kyungnam 작업 때처럼: mapping_list 골격 생성 → AI_INJECT 등 필요한 행 추가 → data_injector.py로 실제 반영 검증
4. **kyungnam 소스 시트 없음 오류 3건** — 상세거래_구축물_차변 / 상세거래_시설장치_대변 / 상세거래_임대보증금(비유동)_대변. 해당 계정의 당기 거래 유무부터 확인 (거래가 없어서 안 만들어진 결과 시트라면 정상, 있는데 안 잡혔다면 버그)
5. **dae_il 이자비용 차이(+90.55%) 원인 확인** — `interest_expense_analysis.py` 실행 자체는 정상화됨(4차 세션에서 SyntaxError 수정). 기대이자 164,940,000원 vs 장부상 1,745,268,147원 차이가 등록 안 된 차입금 때문인지 감사인이 확인 필요

세션 인계 문구: `docs/session_log.md 의 마지막 항목을 읽고, 미해결 이슈부터 이어서 작업해줘.`

---

## 2026-08-12

**완료 작업**: 어제 인계된 1번 항목(kyungnam 27번 Phase 2 '당기감소=당기감가상각비 일치' 현상) 진단 결과, 실제로는 건물↔투자부동산_건물 계정대체가 원인이었고, 그 과정에서 `_depreciation_rollforward()`의 실제 코드 버그(차변/대변 스왑)를 발견·수정 + 계정대체 반영 기능 신규 구현
- **근본 원인 진단**: 사용자가 결과표 I열(`상각누계액_당기감소`)이 라벨과 달리 실제로는 증가값으로 보인다고 지적 → 코드 확인 결과 `_, d_decr = _period_sum(dep_acct)`가 (차변,대변) 튜플에서 **대변(증가)을 감소로 잘못 unpack**하고 있던 실제 버그였음(단순 설계 이슈가 아니었음). 이 때문에 기말잔액 계산식 전체가 왜곡되어 있었음
- **계정대체 반영 설계**: 원가 레벨(취득원가)은 복式부기 특성상 계정대체가 있어도 각 계정의 차/대변 합계에 이미 자동 반영되지만, 감가상각누계액의 '당기증가'는 Phase1에서 감가상각비 상대계정 매칭분만 잡도록 의도적으로 좁혀놨던 터라 대체로 인한 증가(예: 투자부동산_건물감가상각누계액이 대변으로 받는 대체분)가 어디에도 안 잡히고 누락되는 구조였음
- **구현**: `_depreciation_rollforward()` 전면 수정
  - 버그 수정: 상각누계액계정 차변(감소)/대변(증가) 총액을 올바르게 분리
  - `_transfer_amount()` 신규: 전표번호 매칭으로 특정 상대계정과 같은 전표에 있는 금액만 분리 집계 (계정대체 전용, 8번 상대계정분석과 유사한 전표 단위 매칭 방식이나 상대계정을 고정 지정)
  - 취득원가·상각누계액 각각 `당기증가_대체`/`당기증가_기타`/`당기감소_대체`/`당기감소_기타`로 컬럼 분리 (사용자 요청: "증가와 감소에 대체증감을 별도로 표시")
  - `상각누계액_미매칭차이` 컬럼 추가: 감가상각비 매칭분+대체 매칭분으로 설명 안 되는 잔여 대변(증가) 금액을 노출 → `취득원가_수동조정`/`상각누계액_수동조정` 파라미터 컬럼에 감사인이 직접 입력하면 기말잔액 계산에 반영 (사용자 요청: "코딩으로 해결 안 되는 금액은 사용자가 직접 입력")
  - `대체상대계정` 파라미터는 콤마로 복수 지정 가능하도록 설계 — 실제 kyungnam 검증 중 건물↔투자부동산_건물 관계가 단순 1:1이 아니라 투자부동산_건설중인자산(완성대체)까지 얽힌 3계정 구조임이 드러났기 때문 (건물→투자부동산_건물 재분류 + 투자부동산_건설중인자산→투자부동산_건물 완성대체)
- **kyungnam 실 데이터 검증** (`--task 27` 임시 Y 실행 후 다시 N 원복):
  - 취득원가 레벨은 완전히 정합 확인: 건물 감소(1,144,541,982) + 투자부동산_건설중인자산 감소(11,147,763,220) = 투자부동산_건물 증가(12,292,305,202) 정확히 일치
  - 감가상각누계액 레벨은 여전히 잔여차이(`상각누계액_미매칭차이`) 존재 — 원가대체는 한 전표로 일괄기표되지만 상각누계액 대체는 별도/분할 전표로 기표되는 것으로 추정, 정확한 원인은 감사인 확인 필요(설계대로 컬럼에 노출됨)
  - 검증 중 건설중인자산(일반 유형자산)→건물 완성대체 경로도 추가했으나 26년 실적금액은 0(건물 자체 당기 차변이 없음) — 향후 발생 대비 경로만 유지하기로 사용자와 합의
- kyungnam·sejoong task_list의 '감가상각_유형자산롤포워드' 시트에 컬럼 3개 추가(`대체상대계정`/`취득원가_수동조정`/`상각누계액_수동조정`), kyungnam에는 `투자부동산_건설중인자산` 행 신규 추가 + 건물/투자부동산_건물/건설중인자산 행에 실제 대체관계 세팅. sejoong은 계정명이 아직 예시라 대체상대계정은 비워둠(계정명 확정 후 채워야 함)

**변경 파일**:
- `journal_analyzer/main_analyzer.py`: `_depreciation_rollforward()` 차변/대변 버그 수정 + 계정대체(대체/기타 분리, 복수 상대계정, 미매칭차이, 수동조정) 로직 신규, 상단 파라미터 규격 주석 갱신
- `journal_analyzer/kyungnam/task_list_kyungnam.xlsx`: '감가상각_유형자산롤포워드' 시트에 컬럼 3개 추가 + 투자부동산_건설중인자산 행 추가 + 실제 대체관계 세팅
- `journal_analyzer/sejoong/task_list_sejoong.xlsx`: 동일 컬럼 3개 추가(값은 비움)

**미해결 이슈**:
- kyungnam 상각누계액 레벨 미매칭차이(건물 +21,634,715 / 투자부동산_건물 -287,501,957) 원인 미확인 — 감사인이 실제 전표 확인 후 필요 시 `상각누계액_수동조정`에 입력
- 어제 인계 2~5번(27번 파라미터 계정명 정정, sejoong mapping_list, kyungnam 소스 시트 없음 오류 3건, dae_il 이자비용 차이) 전부 그대로 이월

**다음 할 일**:
1. kyungnam 상각누계액 미매칭차이 원인 확인 (실제 전표 조회 필요할 수 있음) 후 수동조정 입력 여부 결정
2. 다른 유형자산 계정들(구축물/기계장치/차량운반구/시설장치/비품/사용권자산)에도 계정대체가 있는지 확인 — 있다면 해당 행에도 `대체상대계정` 채우기
3. (이월) 27번 파라미터 계정명 정정, sejoong mapping_list, kyungnam 소스 시트 없음 오류 3건, dae_il 이자비용 차이

---
