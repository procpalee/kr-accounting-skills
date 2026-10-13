# 이 스킬을 회사 서식으로 바꾸기

남이 만든 서식 스킬을 그대로 쓰지 않고 우리 회사 색·글꼴로 바꾸는 절차다.
**바꾸는 곳은 `scripts/excel_theme.py` 상단 편집 영역(`PALETTES`) 한 곳뿐이다.**
Office.js 쪽 `scripts/theme.json` 과 `references/themes.md` 는 손으로 고치지 않고 `sync_theme.py` 로 다시 만든다.
그래야 클로드 앱·클로드 코드(openpyxl)와 클로드 엑셀(Office.js)이 계속 같은 서식을 낸다.

## 1. 원본을 복사해 새 스킬로 만든다
1. 등록된 원본 스킬 폴더를 작업 폴더로 통째로 복사한다
   (클로드 앱 `/mnt/skills/user/excel-theme/`, 클로드 코드 `~/.claude/skills/excel-theme/`).
   원본이 안 보이면 사용자에게 원본 ZIP 을 첨부해 달라고 한다.
2. 새 이름을 정한다 - **영문 소문자·숫자·하이픈만**(예: `acme-excel-theme`). 한글 이름은 쓸 수 없다.
   폴더명과 `SKILL.md` 의 `name:` 을 같은 이름으로 바꾸고, `description` 첫 문장에 회사명을 넣는다
   (예: "ACME 회계팀 조서 서식을 입힌다."). 원본 스킬과 설명이 같으면 둘 중 무엇을 쓸지 헷갈린다.
3. `SKILL.md` 의 스킬 폴더 경로(`/mnt/skills/user/excel-theme/` 등)도 새 이름으로 바꾼다.

## 2. 편집 영역에서 값만 바꾼다
`PALETTES["default"]` 를 고친다(테마는 default 하나). 사용자 말 → 고칠 키:

| 사용자 말 | 키 (`PALETTES["default"]` 안) |
|---|---|
| 머리글(헤더) 색 | `colors.excel_header_fill` |
| 머리글 글자색 | `colors.excel_header_text` |
| 제목 밴드 색 / 제목 글자색 | `colors.title_bg` / `colors.title_font_color` |
| 합계행 색 | `colors.excel_total_fill` |
| 서브표(두 번째 표) 헤더·합계 색 | `colors.excel_header2_fill` / `colors.excel_header2_text` |
| 표 위아래 굵은 선 색 | `colors.border_out` |
| 표 안쪽 가는 선 색 | `colors.border_in` |
| 입력 셀 색 / 참조 셀 색 | `colors.note` / `colors.linked` |
| 시트 탭 색 | `tabs` (역할 → hex, 예 `{"output":"1F3864", …}`) |
| 엑셀 글꼴 | `font_excel` |

- 색은 `#` 없는 6자리 hex 로 쓴다(`"1F3864"`).
- 말하지 않은 값은 그대로 둔다. 요청이 모호하면(예: "파란색으로") 바꿀 키와 hex 를 먼저 보여 주고 확인받는다.
- **규칙은 바꾸지 않는다** - 계산값은 수식으로, 제목 밴드는 시트명만, 행높이·틀고정 금지, 캡션 40자 등은
  어느 회사에서나 그대로 둔다(`SKILL.md` 2절). 사용자가 명시적으로 요청할 때만 손댄다.

## 3. 동기화하고 검증한다
```bash
python scripts/sync_theme.py          # theme.json · references/themes.md 다시 만들기
python scripts/sync_theme.py --check  # "OK" 가 나와야 한다
```
그다음 두 엔진이 같은 색을 내는지 확인한다.
1. **openpyxl**: `write_table` 로 3열×3행 샘플 표를 만들어 저장하고, 헤더·합계행 셀의 채우기색을 읽어 바꾼 hex 와 같은지 본다.
   `python scripts/theme_lint.py <샘플.xlsx> default` 가 위반 0 이어야 한다(테마 이름은 그대로 default).
2. **Office.js**: `python scripts/make_code.py table '{"n_cols":3,"n_rows":3,"total_rows":[3]}'` 출력에
   바꾼 hex 가 들어 있는지 본다.
3. 바꾼 값 표(항목 · 전 → 후)를 사용자에게 보여 준다.

## 4. 패키징하고 등록한다
- 새 폴더를 ZIP 으로 묶는다(ZIP 안 최상위가 새 스킬 폴더, `__pycache__` 는 뺀다).
- 사용자는 claude.ai **설정 → 사용자 지정(Customize) → 스킬 → 업로드**로 올린다. 같은 계정의 클로드 엑셀에도 그대로 쓰인다.
- 클로드 코드에서 쓰려면 폴더를 `~/.claude/skills/<새 이름>/` 에 둔다.
- 원본 `excel-theme` 을 계속 켜 두면 두 스킬이 같은 요청에 걸릴 수 있다 - 회사 스킬만 쓰려면 원본을 끈다.
