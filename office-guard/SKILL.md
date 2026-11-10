---
name: office-guard
description: 코드가 사람이 고친 Office 문서(xlsx·xlsm·docx·pptx·dsd)를 덮어쓰지 못하게 막는 저장 안전장치. 빌더 실행이 "[저장 거부]"로 멈췄을 때, 빌더를 새로 짤 때 공용 저장 함수(safe_save)를 쓸 때, "덮어써도 돼 / 이 파일 다시 만들어도 돼 / 왜 저장이 안 돼 / 마감했어 / 이제 이 파일이 정본이야" 요청에 사용한다.
---

# office-guard - 사람이 고친 문서를 코드가 덮어쓰지 못하게

설치: 폴더를 넣은 뒤 `settings.json` 에 `PYTHONPATH` 와 훅을 넣어야 작동한다 - [references/install.md](references/install.md).

## 어떻게 지키나
| 층 | 파일 | 맡는 일 |
|---|---|---|
| 파이썬 안 | `runtime/sitecustomize.py` (settings.json `env.PYTHONPATH`) | openpyxl·python-docx·python-pptx·xlsxwriter·dsdxlsx·pandas 저장, `shutil.copy*`·`shutil.move`·`os.replace`·`os.rename`, pywin32 COM 의 `Save`·`SaveAs`·`SaveAs2`·`SaveCopyAs`·`Close(SaveChanges=True)` |
| 도구 호출 | `hooks/office_guard_hook.py` (settings.json PreToolUse·PostToolUse) | Write·Edit, cp·mv·copy·Copy-Item·Move-Item, 리다이렉트, Set-Content·Out-File·tee, COM `.SaveAs("경로")`, 판 이름 사본 |
| 판단 | `office_guard.py` | 위 둘이 함께 쓰는 판단(`judge`)·장부·백업·CLI |

저장해도 되는 경우: 파일이 없음, 임시 폴더·`_backup`·`_xlsync`·`_docsync` 안, 코드가 저장한 내용 그대로, 기준본과 같음, git 에 커밋된 그대로.
같은 파일을 불러와 고쳐 같은 경로에 저장하는 부분 편집은 백업 뒤 허용하되 코드 소유로 치지 않는다 - 그 뒤 빌더로 새로 만들면 여전히 거부된다.
그 밖은 거부한다. 기록이 없는 기존 파일도 사람이 고친 것으로 본다.
파이썬에서 COM 으로 저장할 때도 덮어쓰기 전에 판단한다 - 사람 파일 위 `SaveAs`·`SaveCopyAs` 는 엑셀이 쓰기 전에 거부되고, 연 파일의 `Save` 는 부분 편집으로 백업 뒤 통과한다.

## 복사·이동은 내용의 주인을 바꾸지 않는다
- 코드가 만든 파일을 복사·이동하면 사본도 코드 소유다.
- 사람 파일을 **복사**하면 사본은 그 경로에서만 코드 소유다(장부 키 `<내용>@<경로>`). 사본은 빌더가 덮어쓸 수 있고, 내용이 같은 원본은 계속 사람 파일이다.
- 사람 파일을 **이동**하면(원본이 사라짐) 옮긴 파일은 계속 사람 파일이다.
- `_backup` 에 남기는 백업은 장부에 올리지 않는다.
- 셸 명령의 원천이 변수·와일드카드라 이름을 알 수 없으면 문서로 보고 같은 규칙을 적용한다.

## "[저장 거부]"가 떴을 때
1. 임시 폴더나 다른 이름으로 출력해 결과를 보여 준다. 빌더의 출력 경로를 인자로 받게 고친다.
2. 사람이 고친 내용을 코드에 옮길 때는 `excel-sync` 흐름(diff → 코드 반영 → 임시 재생성 → verify → snap)을 따른다. snap 하면 다시 만들 수 있는 상태가 된다.
3. 몇 칸만 고칠 일이면 그 파일을 불러와 고쳐 같은 경로에 저장한다.
4. 사용자가 그 파일을 덮어써도 된다고 말했을 때만 `python -X utf8 ~/.claude/skills/office-guard/office_guard.py accept "<경로>"` (훅이 승인 창을 띄운다). 승인하면 승인 전 내용이 파일 옆 `_backup` 에 남고 그 위치를 출력한다.

거부를 환경변수·다른 파이썬 옵션으로 피하지 않는다.

## 정본 전환 - 사용자가 말했을 때만
초안 때는 코드가 정본, 마감 때는 사람 파일이 정본이다. 사용자가 "마감했어 / 이제 이 파일이 정본이야"라고 하면 기록한다. 자동으로 전환하지 않는다.
- `python -X utf8 office_guard.py handover "<파일 또는 폴더>" --note "기말 마감"` - 파일마다 한 줄(`_engagement\정본전환.tsv`, `_engagement` 가 없으면 파일 옆 `_xlsync\`). 폴더를 주면 그 안의 지금 있는 문서를 모두 넣는다.
- 전환 뒤에는 코드가 만든 그대로인 파일도 빌더로 다시 저장하지 못한다. `accept` 로도 풀리지 않는다. 임시 폴더 출력과 그 파일을 불러와 몇 칸 고치는 부분 편집은 된다.
- 되돌리기는 사용자가 말했을 때만 `handover "<경로>" --back` (훅이 승인 창을 띄운다).

## 빌더 이력 - 파일 이름은 하나로
빌더가 문서를 저장하면 그때의 코드가 빌더 옆 `_history/<빌더 이름>/` 에 자동으로 쌓이고 `_history/builds.tsv` 에 어느 문서를 만들었는지 남는다(git 저장소·임시 폴더 스크립트는 제외). 같은 폴더의 지역 모듈도 함께 남는다.
- `_v08`, `_전`, `_bak` 를 붙인 빌더 사본을 만들지 않는다. 같은 파일을 고친다. 이미 있는 사본은 그대로 둔다.
- 빌더(.py): 판 이름이 붙었고 같은 폴더에 원본이나 다른 판이 있으면 훅이 `[사본 금지]` 로 막는다. 대상은 이름이 `_build` 로 시작하거나, `_build` 폴더 아래이거나, 새 내용·형제 파일이 Office 문서를 만드는 코드(openpyxl·win32com·docx·pptx 등)인 경우다. `parse_v2.py` 처럼 문서를 만들지 않는 스크립트는 관여하지 않는다.
- 문서(.xlsx 등): `조서_v25.xlsx`·`조서_백업.xlsx` 처럼 구분자 뒤에 판 이름을 붙인 새 문서는, 같은 이름의 앞 판(가장 최근에 고친 것)이 **코드가 만든 그대로**일 때만 막는다. 그때는 앞 판 파일을 고친다(고치기 전 내용은 `_history`·`_backup`).
  앞 판이 사람이 고친 파일이거나 사람 판의 복사본이면 막지 않는다 - 그때는 새 판을 따로 내는 편이 안전하다.
  앞 판을 그대로 두고 새 판을 내야 하면(사용자가 앞 판을 마감했다고 한 경우) 먼저 `handover "<앞 판>"` 으로 정본 전환을 기록한다. `_backup` 등 면제 위치와 읽기만 하는 명령은 관여하지 않는다.
- `python -X utf8 builder_history.py list <빌더.py>` 판 목록, `diff <빌더.py> <판> [<판2>]` 비교, `restore <빌더.py> <판>` 되돌리기(지금 파일을 먼저 이력에 남김), `which <문서>` 이 문서를 만든 판 찾기.

## 빌더에서 직접 쓸 때
```python
sys.path.insert(0, os.path.expanduser("~/.claude/skills/office-guard"))
from office_guard import safe_save
safe_save(wb, out)                              # 사람이 고친 파일이면 OverwriteRefused
safe_save(doc, out, on_conflict="sidecar")      # 사람이 고친 파일이면 옆에 <이름>_재생성.docx
```
Claude Code 밖(엑셀 추가기능 등)에서 도는 파이썬은 PYTHONPATH 가 없으므로 `safe_save` 를 직접 불러야 지켜진다.
PowerShell 등 파이썬 밖에서 COM 으로 저장하는 스크립트는 저장 전에 `python -X utf8 office_guard.py check "<경로>"` (거부면 종료코드 1).
COM 빌더는 `try: ... finally: xl.Quit()` 으로 짠다 - 저장이 거부되어 예외가 나도 엑셀이 남지 않는다.

## 한계
- COM 의 `wb.Save()` 는 이미 연 파일의 부분 편집이라 막지 않고 백업만 한다.
- 파이썬 COM 가드는 pywin32 동적 디스패치(`Dispatch`·`DispatchEx`)만 본다. `gencache.EnsureDispatch` 로 만든 클래스, xlwings(시험하지 않음), PowerShell COM 은 걸리지 않는다 - PowerShell 은 명령에 경로가 보일 때 훅이 본다.
- `shutil.copytree` 로 폴더째 복사한 문서는 판단을 거치지 않는다.
- 다른 프로세스가 만든 `.tmp` 를 문서 이름으로 옮기면 원천을 알 수 없어 코드 소유로 올라간다.
- 저장이 거부되면 빌더가 그 전에 만든 중간 파일은 남는다. `xl.Quit()` 을 `finally` 에 두지 않은 빌더는 엑셀도 남긴다.
- 이미 장부에 올라간 항목은 규칙이 바뀌어도 소급해 내려가지 않는다. 사람 파일이 올라가 있으면 `office_guard.py revoke "<경로>"` 로 내린다(승인 취소에도 쓴다. 내용이 같은 다른 파일의 기록도 함께 내려간다).
- 파이썬이 `open(경로, "wb")` 로 문서를 직접 쓰는 경우, Node 등 다른 언어 빌더는 파이썬 층에 걸리지 않는다(경로가 명령에 보이면 백업).
- 사람이 고쳤는지는 파일 내용의 해시로 본다. 엑셀에서 열어 저장만 해도 바뀐 것으로 잡힌다 - `excel-sync diff` 로 실제 차이를 보고 snap 한다.
