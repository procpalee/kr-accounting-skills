# 설치 (처음 한 번)

필요한 것: 윈도우, 클래식 Outlook(데스크톱판. new Outlook 에는 VBA 가 없음), 파이썬 3.11 이상, `pywin32`(`pip install pywin32`).

## 1. 설정 파일 만들기
```
python -X utf8 scripts/setup.py --root "D:\거래처"
```
- 거래처 루트 아래에 거래처마다 폴더를 하나씩 만든다(`D:\거래처\가나상사`, `D:\거래처\03. 다라정밀`). 앞의 `NN. ` 번호는 비교할 때 뗀다.
- 고객사·협력사처럼 한 단계 묶고 싶으면 `--groups "01. 고객사,02. 협력사"` - 그 폴더 아래에 거래처 폴더를 둔다.
- Outlook 쪽 부모 폴더 이름은 기본 `거래처`(`--outlook-root` 로 바꿈). 기본 사서함 최상위(받은 편지함과 같은 줄)에 생긴다.
- `<루트>\_메일분류\mail_sort.json` 에서 감시 폴더(`watch`), 제외 제목, 단체 메일 기준(`max_recipients`), 우리 쪽 사람(`team`), 수집 시작일(`since`)을 고친다.

## 2. Outlook 에 VBA 넣기
1. Outlook 에서 `Alt+F11`(VBA 편집기).
2. 파일 > 파일 가져오기 → `<루트>\_메일분류\vba\ClientSort.bas`, 다시 가져오기 → `ClientSortWatch.cls`.
3. 왼쪽 프로젝트 창의 `ThisOutlookSession` 을 열고 `ThisOutlookSession_붙여넣기.txt` 내용을 붙여 넣는다(같은 이름 Sub 가 이미 있으면 그 안에 `ClientSort.Start "…"` 한 줄만).
4. 디버그 > 컴파일 → 저장.
5. 매크로 보안: Office 폴더의 `SELFCERT.EXE` 로 개인 인증서를 만들고 VBA 편집기 도구 > 디지털 서명에서 서명한 뒤, Outlook 보안 센터에서 '디지털 서명된 매크로만 허용'. (서명 없이 '모든 매크로 알림'으로 둬도 되지만 켤 때마다 묻는다.)
6. Outlook 을 껐다 켠다. `<루트>\_메일분류\last_run.txt` 가 생기면 시작된 것.

## 3. 버튼 달기 (빠른 실행 도구 모음)
파일 > 옵션 > 빠른 실행 도구 모음 > 명령 선택 '매크로' 에서 세 개를 추가:
| 매크로 | 버튼 이름 |
|---|---|
| `ClientSort_File` | 거래처로 분류 |
| `ClientSort_LinkFolders` | 폴더 연결 맞추기 |
| `ClientSort_Help` | 사용법 |

## 4. 거래처 연결
거래처 폴더를 만든 뒤 [폴더 연결 맞추기]를 누르거나 `python -X utf8 scripts/mail_sync_all.py --link-only`.
Outlook 에 같은 이름 폴더가 있으면 연결하고, 없으면 만든다. 거래처마다 `_mail\mail.json` 이 생기므로 거래처 담당자 주소·도메인을 넣는다(또는 [거래처로 분류]에서 '기억').

## 5. 늘 켜 두기 (선택)
`<루트>\_메일분류\outlook_autostart.ps1` 을 가리키는 바로가기를 `shell:startup` 에 넣는다.
대상: `powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "<루트>\_메일분류\outlook_autostart.ps1"`
바로가기를 '최소화로 실행'으로 두지 않는다 - 그렇게 켜면 Outlook 이 매크로를 불러오지 않는다.

## VBA 를 바꿀 때 (스킬 업데이트 뒤)
`python -X utf8 scripts/setup.py` 로 VBA 파일을 다시 만들고, VBA 편집기에서 모듈 `ClientSort`·클래스 `ClientSortWatch` 를 제거(내보내기 아니요) → 다시 가져오기 → 컴파일 → 서명 → 저장 → Outlook 재시작.
