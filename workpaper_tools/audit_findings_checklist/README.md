# 감리 지적사항 ↔ 계정리스트 대사 체크리스트 생성기

금융감독원·한국공인회계사회의 감리 지적사례(PDF/HWP/HWPX)를 구조화된 DB로 만들고, 회사 계정리스트와
대사해 표준 계정 분류별 위험·확인사항 체크리스트(엑셀)를 만든다. 결과물은 감사조서 작성의 참고자료이며
최종 위험평가는 감사인의 판단이다.

## 실행 순서

`workpaper_tools\audit_findings_checklist` 폴더에서 anaconda 파이썬으로 실행한다.

| 순서 | 명령 | 하는 일 |
|---|---|---|
| 1 | `python main.py run-all` | `input_pdfs`의 새 원문만 추출·분할하고 구조화 예상 비용을 보여준다 (API 호출 없음) |
| 2 | `python main.py run-all --yes --batch` | 구조화까지 실행. `--batch`는 반값이지만 결과까지 수 분~1시간. 급하면 `--batch`를 뺀다 |
| 3 | `python main.py review-export` | 검토용 엑셀 `output\review_날짜.xlsx` 생성 |
| 4 | (엑셀에서 검토·수정 후 저장) | 노란 머리글 열만 수정. 검토상태를 확정/제외로 바꾼다 |
| 5 | `python main.py review-import` | 수정 내용을 DB에 반영 |
| 6 | `python main.py map --accounts <분석결과_회사.xlsx>` | 회사 계정을 표준 분류로 매칭 (`--no-llm`: 외부 전송 없이 사전만) |
| 7 | `python main.py report --company <회사>` | 체크리스트 `output\checklist_회사_날짜.xlsx` 생성 |

6번의 계정리스트는 `input_accounts` 폴더에 **회사명을 파일명으로** 넣는다(예: `input_accounts\samdong.xlsx`). journal_analyzer의 `분석결과_회사.xlsx`를 복사해 이름만 바꾸면 되고, 계정명 열이 있는 다른 xlsx/csv도 된다. 런처는 이 폴더에 파일이 있는 회사만 6·7번에서 보여주며, 파일명을 회사명(`--company`)으로 넘긴다.

`python main.py status`로 언제든 현황(구조화 대기 건수, 검토 상태, 진행 중인 배치)을 볼 수 있다.

새 원문을 `input_pdfs`에 넣고 1~2번을 다시 실행하면 새 파일·새 사례만 처리된다(파일 해시와 사례번호로 판별).
배치가 시간 안에 끝나지 않으면 `python main.py structure --batch`를 다시 실행해 이어받는다.

## 설정 파일 (`config\`)

| 파일 | 내용 |
|---|---|
| `config.yaml` | 모델명, 단가, 경로, 제외 파일, 체크리스트 한 행에 담는 건수 등 |
| `segment_rules.yaml` | 지적사례 분할 규칙(정규식) |
| `account_taxonomy.yaml` | 표준 계정 분류와 키워드 사전, 회사별 예외(`overrides`) |
| `finding_types.yaml` | 지적유형 허용값 |

API 키는 프로젝트 루트 `.env`의 `ANTHROPIC_API_KEY`에서 읽는다.

## 원칙

- 체크리스트의 모든 위험 항목에는 finding_id·출처 파일·페이지·원문 발췌가 붙는다.
- 원문 발췌는 코드가 원문과 대조한다. 발췌가 없거나 원문과 다른 지적사항은 체크리스트에 싣지 않는다.
- "권장 확인사항"만 AI 제안이며 그렇게 표시된다.
- 외부(Claude API)로 보내는 회사 정보는 계정명·계정코드뿐이다. 금액은 읽지 않는다.
- HWP/HWPX는 한컴오피스가 설치된 PC에서만 읽을 수 있다.

## 테스트

`python -m pytest tests -q` (API를 호출하지 않는다)
