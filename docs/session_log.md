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
