---
name: outlook-client-mail
description: 클래식 Outlook 메일을 거래처별로 자동 분류하고(Outlook VBA - 제목의 거래처명, 주소, 같은 대화 순서의 정해진 규칙), 거래처 폴더에 원본(.msg)과 받은 첨부를 저장한 뒤, Claude 가 거래처마다 '메일 히스토리'(사람과 역할, 타임라인, 확인 사항, 챙길 것, 주고받는 패턴)를 원본 링크와 함께 정리한다. "메일 자동 분류 설치해줘 / 거래처별로 메일 정리해줘 / OO 메일 정리해줘 / 메일 히스토리 업데이트 / OO 랑 주고받은 메일 찾아줘 / 그 메일 열어줘 / 거래처 추가했어 / 챙길 것 알려줘" 요청에 사용. 윈도우 + 클래식 Outlook + pywin32 필요.
---

# outlook-client-mail - 거래처별 메일 자동 분류 · 저장 · 메일 히스토리

도구 폴더: 이 스킬의 `scripts/`(실행은 항상 `python -X utf8 scripts/<도구>.py …`).

## 역할 나누기
| 층 | 누가 · 언제 | 하는 일 | 방식 |
|---|---|---|---|
| ① 분류 | Outlook VBA `ClientSort` - 메일이 들어오는 즉시 | 받은 · 보낸 · 감시 폴더의 메일을 거래처 Outlook 폴더로(받음 이동 · 보냄 사본). 못 정하면 그대로 - 사용자가 끌어다 넣거나 [거래처로 분류] | 결정론 |
| ② 저장 | VBA → `mail_sync.py` - 거래처 Outlook 폴더에 메일이 들어오는 즉시 | 원본 `.msg` → `<거래처>\메일\YYMM\`, 색인 `_mail\mail_index.jsonl`, 받은 첨부 → `<거래처>\첨부\` | 결정론 |
| ③ 연결 | VBA → `mail_sync_all.py` - Outlook 을 켤 때 · [폴더 연결 맞추기] | 거래처 폴더 ↔ Outlook 폴더 연결(`mail.json` 의 `outlook.entry_id`), 새 Outlook 폴더 만들기, 거래처 목록(`Client_Map.tsv`) | 결정론 |
| ④ 정리 | 이 스킬(Claude) - 요청할 때 | 메일 히스토리 갱신, 찾기 · 열기, 역할 모르는 주소 확인 | AI |

Claude 는 ①~③의 결과를 읽기만 한다. 분류 규칙을 바꾸고 싶으면 VBA 가 아니라 `mail.json`(주소 · 도메인 · 키워드)과 `mail_sort.json` 을 고친다.

## 폴더
```
<거래처 루트>\
  _메일분류\  mail_sort.json(사용자 설정) · Client_Map.tsv(생성물, 직접 고치지 않음) · Client_Map_추가.tsv(VBA 가 기억한 주소)
              Client_Sort_log.tsv(분류 기록, UTF-16) · last_run.txt · last_sync.txt · vba\ · 사용법.html
  [그룹\]<거래처>\  메일\YYMM\*.msg · 첨부\<메일 단위>\ · 메일 히스토리.md · _mail\(mail.json · mail_index.jsonl · mail_state.json, 숨김)
Outlook 기본 사서함 최상위\<outlook_root(기본 '거래처')>\[그룹\]<거래처>
```
루트 위치는 `%APPDATA%\outlook-client-mail\config.json`(setup.py 가 씀) 또는 환경변수 `OCM_ROOT`.

## 도구
| 도구 | 하는 일 |
|---|---|
| `setup.py --root <폴더> [--outlook-root 거래처] [--groups "a,b"]` | 처음 한 번 - 설정 · VBA 파일 · 사용법 · 자동 시작 스크립트 만들기. 스킬 업데이트 뒤에는 인자 없이 다시(VBA 다시 만들기) |
| `mail_sync_all.py [--link-only \| --status]` | 플래그 없음 = 연결 + 거래처 목록 + 모든 거래처 수집(VBA 가 켤 때), `--link-only` = 연결 + 목록(버튼), **`--status` = 요약만(이 스킬)** |
| `mail_sync_all.py --mark <거래처> <번호>` | 히스토리를 #번호까지 반영했다고 기록 |
| `mail_sync.py <거래처> [--dry-run] [--full] [--since]` | 한 거래처 수집(VBA 가 부름). 이미 모은 메일은 건너뜀(Internet Message-ID). 겹치면 잠금 |
| `outlook_folders.py [--apply]` | 거래처 폴더 ↔ Outlook 폴더 연결 미리보기 · 맞추기 |
| `client_map.py [--show]` | VBA 가 읽는 거래처 목록 만들기. VBA 가 기억한 주소는 거래처 mail.json 에 반영 |
| `mail_search.py <거래처> [키워드…] [--from] [--since] [--until] [--dir 받음\|보냄] [--attach] [--after N] [--body] [--open N]` | 색인에서 찾기(키워드 AND). `--open N` 은 원본을 Outlook 으로 연다 |
| `mail_links.py <거래처> <원고.md> [<출력.md>]` | 원고의 `{#번호}` → 원본 링크 `[#번호](<메일/YYMM/….msg>)`. 출력을 빼면 `메일 히스토리.md` |

## 절차 A - 설치 ("메일 자동 분류 설치해줘")
1. 거래처 루트 위치와 Outlook 부모 폴더 이름, 그룹이 필요한지 사용자에게 묻는다. 정하면 `setup.py --root …`.
2. `references/install.md` 2~3절(VBA 가져오기 · 서명 · 버튼)은 **사용자가 직접** 한다 - 순서를 안내만 한다. 매크로 보안 설정을 Claude 가 바꾸지 않는다.
3. 거래처 폴더를 만들었으면 `mail_sync_all.py --link-only` 결과를 보여 준다. 새로 생긴 `_mail\mail.json` 에 넣을 담당자 주소 · 도메인을 사용자에게 받는다.
4. `mail_sort.json` 의 `watch`(팀 공용 폴더 등), `team`(우리 쪽 사람), `since`(수집 시작일)를 사용자와 정한다.

## 절차 B - 메일 히스토리 ("OO 메일 정리해줘 / 히스토리 업데이트")
1. `mail_sync_all.py --status` - 거래처별 '정리할 메일 #a ~ #b', 새 첨부, 대용량 링크, 역할 모르는 보낸 사람. **수집 · 연결은 하지 않는다**(VBA 몫).
   새 메일이 VBA 로 안 들어온 것 같으면(Outlook 이 꺼져 있었음 등) 사용자에게 Outlook 을 켜 달라고 하거나, 동의를 받아 `mail_sync.py <거래처>` 를 직접 돌린다.
2. `mail_search.py <거래처> --after <a-1> --body --limit 200` 으로 새 메일을 읽는다. 출력의 '링크:' 를 근거로 그대로 쓴다.
3. 히스토리 문서(`references/history-format.md` 형식)를 고친다. 원고는 `{#번호}` 로 근거를 달고 `mail_links.py` 로 링크를 만든다. 문서가 없으면 새로 만든다.
   자동 알림 · 단순 전달처럼 반영할 게 없는 메일은 건너뛰되 번호는 넘긴다. 사람이 고친 상태 · 기한 · 메모는 덮어쓰지 않는다.
4. 끝나면 `mail_sync_all.py --mark <거래처> <b>`.
5. **답이 안 보이는 요청은 묻는다**: 메일로 회신이 없는 질문 · 요청은 바로 '미결'로 적지 말고 전화 · 대면으로 해결됐는지 사용자에게 묻고 그 답으로 상태를 쓴다.
6. **역할 모르는 보낸 사람**: 히스토리 1절에서 거래처 쪽 사람으로 확인되면 보고에 'mail.json 거래처 주소에 넣을지' 묻는다(넣으면 VBA 자동 분류에도 쓰임). 동의하면 `company.addresses` 에 넣고 `client_map.py`.
7. **보고**: 새 메일 수, 새 '챙길 것', 완료 후보, 물어볼 것(회신 없는 요청 · 역할 모르는 주소), 7일 안에 만료되는 대용량 링크.

## 절차 C - 찾기 · 열기 ("OO 랑 주고받은 견적 메일 찾아줘 / 그 메일 열어줘")
`mail_search.py <거래처> <키워드…>` → 목록을 보여 주고, 열어 달라면 `--open N`. 어느 거래처인지 모르면 거래처 이름을 묻거나 여러 거래처를 차례로 찾는다.

## 지키는 것
- **메일함은 읽기만**: Claude 는 메일을 보내거나 지우거나 옮기지 않는다(옮기기는 VBA 와 사용자). 도구도 Outlook 에서 원본을 바꾸지 않는다(사본 저장만).
- 메일 원본 · 색인 · 히스토리에는 거래처 정보가 들어간다 - 거래처 루트를 공개 저장소 · 클라우드 공유 폴더에 두지 않는다.
- `Client_Map.tsv` 는 직접 고치지 않는다(다음 연결 때 다시 만들어짐). 고칠 곳은 `mail.json` · `mail_sort.json` · `Client_Map_추가.tsv`.
- 클래식 Outlook 이 켜져 있을 때만 분류 · 저장이 돈다(new Outlook 에는 VBA 가 없다 - 켜 두면 결과 폴더는 new Outlook 에서도 보인다). 메일함에 남은 기간보다 오래된 메일은 모을 수 없다.
