---
name: excel-theme-officejs
description: 열려 있는 엑셀 시트의 표와 시트를 사내 excel-theme 서식으로 꾸민다. 테마는 default(정산표·결산), audit(감사조서), procpa(브랜드 블루), dcf-valuation(DCF·평가) 4종. "표 서식 입혀줘", "테마 적용", "양식 통일", "정산표처럼", "조서 양식으로", 합계·소계·섹션행 강조, 감사조서 헤더, 제목 밴드, 표 캡션, 시트 탭 색을 요청할 때 쓴다. 동봉 스크립트가 같은 입력이면 항상 같은 Office.js 코드를 만들고, 그 코드를 수정 없이 실행한다.
---

# excel-theme (Office.js)

엑셀 표·시트를 사내 표준 서식으로 맞추는 스킬이다.
**서식 코드를 직접 짜지 않는다.** `scripts/make_code.py` 가 출력한 Office.js 코드를 **한 글자도 바꾸지 않고** 실행한다.
그래야 누가 언제 만들어도 서식이 똑같다(색·테두리·숫자서식 규칙은 스크립트 안에 고정돼 있다).

## 1. 테마 고르기
| 상황 | theme |
|---|---|
| 감사조서(계정조서·중간/기말감사) | `audit` |
| 브랜드 산출물(고객 제출용 보고서 표) | `procpa` |
| DCF·가치평가·재무모델 | `dcf-valuation` |
| 그 외 정산표·결산·일반 표 (기본) | `default` |

사용자가 테마를 말하지 않으면 위 표로 고르고, 애매하면 `default`. 이 4종 외의 테마는 없다.

## 2. 작업 순서
1. **값·수식부터 넣는다.** 합계·소계·차이·비율·건수처럼 계산되는 값은 반드시 엑셀 수식(SUM, SUMIFS…)으로 넣는다.
   숫자를 계산해서 값으로 넣지 않는다. 원자료·판단 파라미터·텍스트만 값이다.
2. **표 위치를 규약대로 잡는다.** A열과 1행은 비운다. 표는 B열에서 시작한다.
   - 일반 시트: B2 제목 밴드 → 4행 헤더
   - 감사조서: B2:G3 조서 헤더 → 5행부터 본문
   - 캡션(표 소제목)을 두면 캡션 행 +2 가 헤더 행(바로 아래 한 줄은 비움)
3. **시트 요소 코드**를 만들어 실행한다(필요할 때): `python scripts/make_code.py parts '<JSON>'`
4. **표 코드**를 만들어 실행한다: `python scripts/make_code.py table '<JSON>'`
   - 3번 출력의 `// next_header_row` 값을 표의 `header_row` 로 쓴다.
   - 한 시트에 표가 여러 개면 표마다 한 번씩 실행한다. 두 번째 표부터 구분이 필요하면 `"secondary": true`.
5. 출력된 코드를 **그대로** Office.js로 실행한다. 코드 앞의 `// WARNING:` 줄은 사용자에게 알린다.
   - 실행 도구가 `context` 를 이미 주는 방식이면 JSON 에 `"body_only": true` 를 넣어 다시 만든다.
6. 결과를 확인한다: 헤더·합계행 몇 칸의 채우기색·굵기·테두리를 읽어 보고, 수식 오류(#REF! 등)가 없는지 본다.

스크립트 경로는 이 스킬 폴더 안의 `scripts/make_code.py` 다(예: `/mnt/skills/user/excel-theme-officejs/scripts/make_code.py`).
표준 라이브러리만 쓰므로 설치할 패키지가 없다.

## 3. make_code.py 입력

### table
```json
{"theme":"default", "n_cols":4, "n_rows":6,
 "currency_cols":[1,2], "percent_cols":[3],
 "section_rows":[1], "subtotal_rows":[4], "total_rows":[6],
 "input_cells":["C6:D7"], "linked_cells":["C9:D9"], "todo_cells":["E6"]}
```
| 키 | 뜻 |
|---|---|
| `n_cols`, `n_rows` | 헤더 열 수, 데이터 행 수(헤더 제외) — 필수 |
| `header_row` | 헤더 행 번호(기본 4) |
| `start_col` | 표 시작 열 번호(기본 2 = B) |
| `title_cell` | `"auto"`(기본) = B2 제목 밴드, audit 은 제목 밴드 없음 · `""` = 제목 서식 안 함 |
| `currency_cols` | 통화 열(0-based, 첫 데이터열=0). 원 단위 회계서식(`accounting`). 백만원은 `number_cols` 에 `million_won` |
| `percent_cols` | 비율 열(0-based) |
| `number_format_cols` | 열별 서식 직접 지정 `{"C":"accounting"}` 또는 `{"1":"multiple"}` |
| `section_rows` / `subtotal_rows` / `total_rows` | 데이터 1-based 행 번호(헤더 바로 아래 = 1) |
| `input_cells` / `linked_cells` / `todo_cells` | 입력(크림) · 참조(연회색) · 미입수(노랑) 범위 |
| `secondary` | 서브표(2차 표) 헤더·합계색 |
| `autofit` | 열너비 자동 맞춤(기본 true, 상한 30자) |
| `sheet_name` | 대상 시트(생략 시 활성 시트) |

숫자서식 키: `accounting`, `million_won`, `million`, `billion`, `thousands`, `currency_won`, `percent_acct`, `percent`, `decimal`, `multiple`, `change`, `bp`, `date`.

### parts
```json
{"theme":"audit", "parts":[
  {"type":"audit_header", "meta":{"회사명":"A사","조서번호":"6000A-10","결산일":"2026-12-31","조서명":"매출채권","작성자":"홍길동"}},
  {"type":"section_bar", "row":6, "text":"1. 총괄표", "last_col":8},
  {"type":"caption", "cell":"B8", "text":"표1. 연령별 잔액"},
  {"type":"tab", "role":"output"}]}
```
| type | 내용 |
|---|---|
| `title_band` | B2 제목 밴드(`last_col` 까지 색). **내용은 시트명만** — `text` 를 주지 않으면 시트명(앞 순번 제거)이 들어간다. audit 에는 쓰지 않는다 |
| `audit_header` | 감사조서 헤더 6항목(회사명·조서번호·시트명 / 결산일·조서명·작성자), B2:G3. 시트명은 생략 시 자동 |
| `section_bar` | 섹션 소제목 밴드(`row`, `text`, `last_col`) |
| `caption` | 표 소제목(`cell`, `text`) — 짧은 명사구만, 40자 이내 |
| `tab` | 시트 탭 색 역할: `guide`(안내·표지, 색 없음) · `output`(결과·총괄) · `calc`(계산) · `input`(작성자 입력) · `pbc`(회사 제공 자료) · `raw`(원본·참고) |

## 4. 스크립트가 대신 못 하는 규칙 (직접 지킨다)
- **계산값은 수식으로.** 표의 합계·소계·증감·비율은 SUM/SUMIFS/ROUND 등 수식.
- **제목 밴드에는 시트명만.** 회사명·기준일·설명을 제목에 붙이지 않는다. 설명이 필요하면 캡션이나 비고 열에.
- **캡션은 짧은 제목만**(40자 이내). 계산 방법·근거는 표의 비고 열에 쓴다.
- **긴 대시(— – ―) 대신 하이픈(-)** 을 쓴다. 스크립트가 표·제목의 텍스트는 바꾸지만, 수식 안 문자열과 시트명은 처음부터 `-` 로 쓴다.
- **행높이·틀고정은 건드리지 않는다.** 긴 문장은 줄바꿈·열너비·병합으로 처리.
- **탭 색은 역할로만.** 한 파일에서 같은 역할은 같은 색. 진행 상태 표시에 쓰지 않는다.
  시트 순서: 안내 → 결과 → 계산 → 입력 → 원본.
- **감사조서 시트 이름**은 `10 총괄표`, `20 세부Test` 처럼 10 단위 순번 + 이름, 조서번호는 `6000A-10` 식으로 맞춘다.
- **기존 표를 고칠 때는** 사용자가 말한 범위만 서식한다. 코드를 실행하기 전에 대상 범위를 사용자에게 알린다.

## 5. 스크립트를 실행할 수 없을 때
코드 실행 환경이 없으면 `references/themes.md`(테마별 색·숫자서식·탭 색)와 `references/rules.md`(테두리·합계행·정렬 규칙)를 읽고 같은 규칙으로 직접 서식한다.
이 경우 결과가 스크립트와 조금 다를 수 있다고 사용자에게 알린다.
