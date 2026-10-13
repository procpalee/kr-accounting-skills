# 클로드 엑셀 - 열려 있는 통합문서 서식 (Office.js)

**서식 코드를 직접 짜지 않는다.** `scripts/make_code.py` 가 출력한 Office.js 코드를 **한 글자도 바꾸지 않고** 실행한다.
그래야 누가 언제 만들어도 서식이 똑같다(색·테두리·숫자서식 규칙은 스크립트 안에 고정돼 있고,
클로드 앱·클로드 코드의 openpyxl 헬퍼와 같은 결과를 낸다).

## 1. 작업 순서
1. **값·수식부터 넣는다.** 계산되는 값은 반드시 엑셀 수식으로.
2. **표 위치를 규약대로 잡는다.** A열·1행 비움, 표는 B열. B2 제목 밴드 → 4행 헤더.
   캡션을 두면 캡션 행 +2 가 헤더 행.
3. **시트 요소 코드**를 만들어 실행한다(필요할 때): `python scripts/make_code.py parts '<JSON>'`
4. **표 코드**를 만들어 실행한다: `python scripts/make_code.py table '<JSON>'`
   - 3번 출력의 `// next_header_row` 값을 표의 `header_row` 로 쓴다.
   - 한 시트에 표가 여러 개면 표마다 한 번씩. 두 번째 표부터 구분이 필요하면 `"secondary": true`.
   - 숫자서식 열은 **열 문자**로 적는다(`"currency_cols":["C","D"]`). 번호로 적으면 한 칸 밀리기 쉽다.
5. 출력된 코드를 **그대로** Office.js로 실행한다. 코드 앞의 `// WARNING:` 줄은 사용자에게 알린다.
   - 실행이 실패하면 코드를 고쳐서 다시 돌리지 않는다. 오류 메시지를 사용자에게 그대로 알린다.
   - 실행 도구가 `context` 를 이미 주는 방식이면 JSON 에 `"body_only": true` 를 넣어 다시 만든다.
6. 결과를 확인한다: 헤더·합계행 몇 칸의 채우기색·굵기·테두리를 읽어 보고, 수식 오류(#REF! 등)가 없는지 본다.

## 2. make_code.py 입력
테마는 `default` 하나라 `theme` 은 생략한다(다른 이름을 넣으면 오류).

### table
```json
{"n_cols":4, "n_rows":6,
 "currency_cols":["C","D"], "percent_cols":["E"],
 "section_rows":[1], "subtotal_rows":[4], "total_rows":[6],
 "input_cells":["C6:D7"], "linked_cells":["C9:D9"], "todo_cells":["E6"]}
```
| 키 | 뜻 |
|---|---|
| `n_cols`, `n_rows` | 헤더 열 수, 데이터 행 수(헤더 제외) - 필수 |
| `header_row` | 헤더 행 번호(기본 4) |
| `start_col` | 표 시작 열 번호(기본 2 = B) |
| `title_cell` | `"auto"`(기본) = B2 제목 밴드 · `""` = 제목 서식 안 함 |
| `currency_cols` | 통화 열(열 문자 권장). 원 단위 회계서식. 백만원은 `number_format_cols` 에 `million_won` |
| `percent_cols` | 비율 열(열 문자 권장) |
| `number_format_cols` | 열별 서식 직접 지정 `{"C":"accounting"}` |
| `section_rows` / `subtotal_rows` / `total_rows` | 데이터 1-based 행 번호(헤더 바로 아래 = 1) |
| `input_cells` / `linked_cells` / `todo_cells` | 입력(크림) · 참조(연회색) · 미입수(노랑) 범위 |
| `secondary` | 서브표(2차 표) 헤더·합계색 |
| `autofit` | 열너비 자동 맞춤(기본 true, 상한 30자) |
| `sheet_name` | 대상 시트(생략 시 활성 시트) |

숫자서식 키: `accounting`, `million_won`, `million`, `billion`, `thousands`, `currency_won`, `percent_acct`, `percent`, `decimal`, `multiple`, `change`, `bp`, `date`.

### parts
```json
{"parts":[
  {"type":"title_band", "last_col":8},
  {"type":"section_bar", "row":6, "text":"1. 총괄표", "last_col":8},
  {"type":"caption", "cell":"B8", "text":"표1. 연령별 잔액"},
  {"type":"tab", "role":"output"}]}
```
| type | 내용 |
|---|---|
| `title_band` | B2 제목 밴드(`last_col` 까지 색). 내용은 시트명만(`text` 생략 시 자동) |
| `section_bar` | 섹션 소제목 밴드(`row`, `text`, `last_col`) |
| `caption` | 표 소제목(`cell`, `text`) - 40자 이내 |
| `tab` | 시트 탭 색 역할: `guide`(색 없음) · `output` · `calc` · `input` · `pbc` · `raw` |


## 3. 스크립트를 실행할 수 없을 때
클로드 엑셀에서는 스킬 파일을 파이썬으로 실행하지 못하는 경우가 있다(스킬 파일이 파이썬 실행 환경과 분리돼 있다).
이때 **스크립트 내용을 손으로 옮겨 적어 다시 만들지 않는다** - 한 글자만 바뀌어도(예: 글꼴명 오타) 스크립트를 쓰는 의미가 없다.
대신 `references/themes.md`(색·숫자서식·탭 색)와 `references/rules.md`(테두리·합계행·정렬 규칙)를 읽고
같은 규칙으로 직접 서식하며, 결과가 스크립트와 조금 다를 수 있다고 사용자에게 알린다.
