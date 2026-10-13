---
name: email-draft
description: 업무 메일 초안 작성. "메일 써줘 / 초안 만들어줘 / 보내는 글 작성해줘" 요청 시 표준 본문 구조로 초안을 만들고, 확정되면 .eml 파일(더블클릭 → Outlook 보내지 않은 새 메일)로 저장한다.
---

# email-draft — 메일 초안 표준

업무 메일 초안을 **항상 같은 구조·톤**으로 만들기 위한 스킬. 산출 형식은 **.eml Outlook 초안**
(X-Unsent 헤더 — 더블클릭하면 작성창으로 열림, 사용자 확정 2026-07). 본문 디자인은 **①안 미니멀**
(굵게 강조만·얇은 테두리 표 — 수신자 클라이언트 어디서든 안 깨짐).

## 절차

1. `templates/`에서 메일 유형 골격을 고른다: `자료요청` / `결과송부` / `검토의견`.
   없는 유형은 아래 '본문 표준 구조'만 지켜 작성.
2. **본문을 채팅에 먼저 보여주고 사용자 확인**을 받는다(제목·수신자·첨부 포함).
3. 확정되면 `email_draft.py`로 .eml 생성:
   ```python
   import sys; sys.path.insert(0, r"C:\Users\PC\.claude\skills\email-draft")
   from email_draft import p, bullets, table, section, assemble_body, build_email
   body = assemble_body(greeting="안녕하세요, OOO 님. 이재현입니다.", blocks=[...])
   build_email(subject="[회사명] 주제 — ...의 건", body_html=body,
               attachments=[r"...\파일.xlsx"], out_path=r"...\메일초안_주제_YYYYMMDD.eml")
   ```
4. 저장 위치는 **해당 프로젝트 폴더**(사용자가 지정한 곳). 파일명: `메일초안_{주제}_{YYYYMMDD}.eml`.
5. 메일 발송은 항상 **사용자가 직접** 한다(초안까지만 만든다).

## 본문 표준 구조 (assemble_body가 이 순서를 강제)

1. **인사** — "안녕하세요, {수신자} 님. 이재현입니다."
2. **배경** 1~2문장 — 무슨 건인지
3. **핵심 내용** — 번호 소제목(`section`) + 불릿(`bullets`)/요약 표(`table`)
4. **요청·다음 단계** — 상대가 할 일을 명확히
5. **맺음** — "감사합니다." + "이재현 드림"
- 제목 형식: `[회사명] 주제 — …의 건`
- 첨부가 있으면 본문에 `attachments_note`로 파일명 표기

## 스타일 규칙 (①안 미니멀, 2026-09-09 확정)

- 폰트는 **Noto Sans KR**(대체 맑은 고딕), 크기 **12**(`font-size:12pt` — Outlook 글꼴 크기 12와 같다),
  글자색은 **검정(#000000)만**. 본문·표 모두 같은 크기·색을 쓴다(`FONT_SIZE`·`TEXT_COLOR` 상수).
- 인라인 스타일만. 외부 CSS·웹폰트·이미지 금지 (Outlook/Gmail 호환)
- 강조는 `<b>`만, 색 강조는 쓰지 않는다
- 표: 테두리 1px `#999`, 헤더만 `#F2F2F2` — `table()` 헬퍼가 처리

### Outlook 서명과 충돌하지 않게 하는 규칙 (필수)

`.eml` 초안을 열면 Outlook이 **새 메시지 서명을 자동으로 끼워 넣는다.** 두 가지를 지킨다.

1. **본문에 맺음말·서명을 쓰지 않는다.** `assemble_body`의 `closing`·`sender` 기본값은 `None`이다.
   사용자의 Outlook 서명 첫 줄이 이미 "감사합니다. / 이재현 드림."이라 본문에 또 쓰면 두 번 나온다.
2. **본문 전체를 1칸 표로 감싼다.** Outlook은 서명을 "본문 첫 블록 다음"에 넣어서, 문단을 그대로 나열하면
   인사말과 본문 사이가 갈라진다. `assemble_body`가 완전한 HTML 문서 + `<table><tr><td>` 한 칸으로
   감싸고 맨 끝에 빈 문단을 두므로, 삽입 자리가 본문 바깥(맨 끝)이 된다.

그래도 서명이 중간에 들어가면 Outlook 설정에서 **새 메시지 서명 자동 삽입을 끄고**
본문 끝에 서명 HTML을 직접 넣는다(사용자 서명 파일: `%APPDATA%\Microsoft\Signatures`).

## 파일

- `email_draft.py` — 본문 블록 헬퍼(`p/section/bullets/table/assemble_body`) + `build_email`(.eml)
- `templates/자료요청.md` · `결과송부.md` · `검토의견.md` — 유형별 본문 골격
- 정리 페이지가 필요하면 → `html-theme` 스킬 (별도)
