# office-guard 설치

스킬 폴더를 넣는 것만으로는 작동하지 않는다. Claude Code 설정(`settings.json`)에 두 가지를 넣어야 한다.

| 넣는 것 | 하는 일 |
|---|---|
| `env.PYTHONPATH` | 파이썬이 시작할 때 `runtime/sitecustomize.py`를 읽어 문서 저장을 판단에 태운다 |
| `hooks` (PreToolUse, PostToolUse) | Write, Edit와 셸의 복사, 이동, 리다이렉트를 판단에 태운다 |

Claude에게 "office-guard 설치해 줘. references/install.md 대로"라고 하면 아래 순서를 대신 한다.

## 필요한 것
- 파이썬 3.11 이상. 표준 라이브러리만 쓴다.
- 엑셀 COM 저장까지 지키려면 윈도우, 데스크톱 엑셀, `pywin32`.
- `excel-sync`, `doc-sync`를 함께 쓰면 기준본과 같은 파일은 다시 만들 수 있다고 판단한다. 없어도 작동한다.

## 1. 폴더 두기
`~/.claude/skills/office-guard/`에 압축을 푼다. 다른 곳에 두면 환경변수 `OFFICE_GUARD_HOME`에 그 폴더를 적는다.

## 2. 설정 파일에 넣기
`~/.claude/settings.json`을 먼저 읽고 **기존 내용에 합친다**. 파일을 통째로 바꾸지 않는다.
- `env.PYTHONPATH`가 이미 있으면 윈도우는 `;`로, 맥과 리눅스는 `:`로 이어 붙인다.
- `hooks.PreToolUse`, `hooks.PostToolUse` 배열에 이미 다른 훅이 있으면 항목을 하나 더한다.
- 경로는 절대 경로로 쓴다. `<홈>`은 자기 홈 폴더로 바꾼다(윈도우 예: `C:/Users/hong`).
- `python`이 여러 개면 훅 명령에 파이썬 실행 파일의 전체 경로를 쓴다.

```json
{
  "env": {
    "PYTHONPATH": "<홈>/.claude/skills/office-guard/runtime"
  },
  "hooks": {
    "PreToolUse": [
      {
        "matcher": "Bash|PowerShell|Write|Edit",
        "hooks": [
          { "type": "command", "command": "python \"<홈>/.claude/skills/office-guard/hooks/office_guard_hook.py\"", "timeout": 15 }
        ]
      }
    ],
    "PostToolUse": [
      {
        "matcher": "Bash|PowerShell",
        "hooks": [
          { "type": "command", "command": "python \"<홈>/.claude/skills/office-guard/hooks/office_guard_hook.py\"", "timeout": 15 }
        ]
      }
    ]
  }
}
```

설정은 새 세션부터 적용된다.

## 3. 확인
임시 폴더는 검사에서 빠지므로 작업 폴더 아래에 시험 폴더를 만들어 확인한다.

1. 엑셀에서 빈 파일을 하나 저장한다(예: `시험\사람파일.xlsx`).
2. `python -X utf8 ~/.claude/skills/office-guard/office_guard.py check "시험\사람파일.xlsx"`를 실행한다. 거부 대상이라고 나오면 판단이 작동한다.
3. Claude에게 그 파일을 openpyxl로 새로 만들어 같은 이름으로 저장해 달라고 한다. `[저장 거부]`로 멈추고 파일이 그대로면 파이썬 층이 작동한다.
4. Claude에게 다른 엑셀 파일을 그 파일 위로 복사해 달라고 한다. 훅이 막으면 도구 호출 층이 작동한다.

## 기록이 남는 곳
| 위치 | 내용 |
|---|---|
| `~/.claude/office-guard/` | 코드가 저장한 내용의 장부(파일 내용의 해시만 남는다). 환경변수 `OFFICE_GUARD_STATE`로 바꿀 수 있다 |
| 문서 옆 `_backup/` | 부분 편집, 승인, 쓰기 명령 전에 남기는 사본 |
| 빌더 옆 `_history/` | 문서를 저장한 빌더 코드의 판 |
| `_engagement/정본전환.tsv` 또는 문서 옆 `_xlsync/` | 정본 전환 기록 |

## 끄기와 지우기
`settings.json`에서 위 `PYTHONPATH`와 훅 항목 두 개를 지우면 꺼진다. 장부 폴더와 `_backup`, `_history`는 지워도 문서에는 영향이 없다. 장부를 지우면 코드가 만든 파일도 처음 한 번은 사람 파일로 보고 멈춘다.
