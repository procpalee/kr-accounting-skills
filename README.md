# kr-accounting-skills

한국 회계·재무 실무자를 위한 **Claude 스킬 모음**입니다. 결산 파일·명세·보고서처럼 매번 반복하는 일을
Claude Code 에 맡길 때 쓰는 스킬을 모았습니다. 더존 원장, 한국식 표 서식, 한국어 업무 메일처럼 국내 실무 환경을 기준으로 만들었습니다.
[procpa.co.kr](https://procpa.co.kr/tools)에서 소개·배포합니다.

원칙은 하나입니다. **반복은 코드, 해석은 Claude, 결정은 사람.**
1원까지 같아야 하는 계산·대사·서식은 스크립트로 고정하고, Claude 는 그 스크립트를 쓰고 돌리고 결과를 설명하며,
판단과 확정은 사람이 합니다.

## 스킬 목록

### 업무 자동화 만들기

| 스킬 | 하는 일 | 다운로드 |
|---|---|---|
| [work-skill-creator](work-skill-creator/) | 내 반복 업무(분기 결산 파일 · 명세 · 보고서)를 **업무별 스킬로 만드는 스킬**. 지난 기간 결과를 코드로 다시 만들어 사람이 만든 완성본과 칸마다 비교해 차이 0 을 확인한 뒤 스킬로 저장하고, 다음 기간에 그대로 돌린다 | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/work-skill-creator.zip) |

### 결산 · 장부

| 스킬 | 하는 일 | 다운로드 |
|---|---|---|
| [ledger-merge](ledger-merge/) | 더존에서 내려받은 계정별원장(계정마다 시트 하나)을 시트 하나로 합치고, 계정별 차변 · 대변 합이 원장 [누계]와 맞는지 확인 시트에 적는다 | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/ledger-merge.zip) |
| [excel-sync](excel-sync/) | 엑셀에서 사람이 직접 고친 내용을 기준본과 비교해(값 · 수식 · 서식 · 행열 삽입 · 병합 · 유효성 …) 코드에 반영하고, 다시 만든 파일과 차이 0 을 확인한다 | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/excel-sync.zip) |

### 문서 서식 · 커뮤니케이션

| 스킬 | 하는 일 | 다운로드 |
|---|---|---|
| [excel-theme](excel-theme/) | 엑셀(.xlsx) 생성 · 편집 시 일관된 색 · 폰트 · 표 서식 테마(default · procpa · dcf-valuation) | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/excel-theme.zip) |
| [word-theme](word-theme/) | 워드(.docx) 서식 테마(default · procpa) + 문서 유형별 제목부 프리셋(보고서 · 메모 · 제안서 · 의견서 · 조서) | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/word-theme.zip) |
| [html-theme](html-theme/) | 정리 · 비교 · 검토 결과를 보고서 문서처럼 보이는 HTML 페이지로 | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/html-theme.zip) |
| [email-draft](email-draft/) | 업무 메일 초안(.eml)까지만 만들기 - 발송은 사람이 직접 | [zip](https://github.com/procpalee/kr-accounting-skills/releases/latest/download/email-draft.zip) |

ledger-merge 는 excel-theme 를 함께 설치해야 합니다.

## 설치

1. 위 표의 zip(항상 최신 릴리스)을 받아 압축을 풉니다.
2. 스킬 폴더를 Claude Code 스킬 디렉터리에 넣습니다.
   - 개인용: `~/.claude/skills/<스킬이름>/`
   - 프로젝트용: `<프로젝트>/.claude/skills/<스킬이름>/`
3. 새 세션에서 말로 부릅니다. 예: "대손충당금 결산 파일을 매 분기 자동으로 만들고 싶어. 이 업무 스킬로 만들어 줘", "이 계정별원장 시트 하나로 합쳐 줘".

필요한 것: 파이썬 3.11 이상, `openpyxl`. 엑셀 재계산 · 직접 조작이 들어가는 스킬은 윈도우 + 데스크톱 엑셀과 `pywin32`.

각 스킬의 자세한 사용법은 [procpa.co.kr/tools](https://procpa.co.kr/tools)를 참고하세요.
스킬 개념이 처음이라면 → [스킬이란?](https://procpa.co.kr/guide/what-is-skill)

## 고객 자료 주의

스킬은 내 PC 에서 내 파일을 다룹니다. 업무 스킬을 만들면 그 안에 고객사 규칙 · 판단이 들어가므로
**고객 정보가 들어간 스킬은 그 업무 폴더 안에만 두고** 공개 저장소 · 외부 서비스에 올리지 마세요.

## 릴리스 방식

`v*` 태그를 push 하면 GitHub Actions 가 **SKILL.md 가 있는 최상위 폴더마다 zip 을 만들어 하나의 릴리스에 첨부**합니다.
`releases/latest/download/<스킬이름>.zip` 은 항상 최신 버전을 가리킵니다.

```bash
git tag v2026.10.13 && git push origin v2026.10.13
```

## 문의

버그 제보 · 기능 제안: [Issues](https://github.com/procpalee/kr-accounting-skills/issues) 또는 [procpa.co.kr/contact](https://procpa.co.kr/contact?type=tools)

## License

MIT
