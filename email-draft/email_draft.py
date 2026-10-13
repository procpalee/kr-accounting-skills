# -*- coding: utf-8 -*-
"""
email-draft — .eml Outlook 초안 생성기 (①안 미니멀, 사용자 확정 2026-07)

더블클릭하면 Outlook이 '보내지 않은 새 메일(작성창)'로 여는 .eml을 만든다
(X-Unsent: 1 헤더 — opinion-letter 스킬과 동일 방식).

본문 스타일 원칙(이메일 클라이언트 호환 최우선):
  - 인라인 스타일만 사용, 외부 CSS·웹폰트 금지
  - 강조는 <b>만, 색은 쓰지 않음(수신자 클라이언트 기본값 상속)
  - 표는 얇은 테두리(1px #999) + 헤더 옅은 회색(#F2F2F2)만

사용 예:
    import sys; sys.path.insert(0, r"C:\\Users\\PC\\.claude\\skills\\email-draft")
    from email_draft import p, bullets, table, section, assemble_body, build_email

    body = assemble_body(
        greeting="안녕하세요, OO팀 OOO 님. 이재현입니다.",
        blocks=[
            p("결산개선 과제로 진행한 <b>계정과목 체계 통일안</b>을 송부드립니다."),
            section("1. 주요 내용"),
            bullets(["ERP 입력계정 775건 → 정산표·공시 계정 매핑 완료",
                     "미사용 계정 53건 정리, 명칭 표기 규칙 통일"]),
            table(["구분", "건수", "비고"],
                  [["자동 매핑", "592", "수식 기반"], ["판단 매핑", "183", "근거 별도 표시"]]),
            section("2. 요청 사항"),
            p("첨부 파일의 '판단매핑' 시트 검토 후 회신 부탁드립니다."),
        ],
    )
    build_email(subject="[OO사] 결산개선 — 계정과목 통일안 송부의 건",
                body_html=body, attachments=[r"...\\COA_통일안_v8.xlsx"],
                out_path=r"...\\메일초안_COA통일안_20260722.eml")
"""
import os
import re
import mimetypes
from email.message import EmailMessage
from email.utils import formatdate
from email.generator import BytesGenerator

FONT = "font-family:'Noto Sans KR','Malgun Gothic','맑은 고딕',sans-serif;"
FONT_SIZE = "12pt"                 # Outlook 글꼴 크기 12 (사용자 확정 2026-09-09)
TEXT_COLOR = "#000000"             # 글자색은 검정만
BODY_STYLE = f"{FONT} font-size:{FONT_SIZE}; line-height:1.7; color:{TEXT_COLOR};"
_TD = "border:1px solid #999; padding:4px 10px;"

SENDER = None       # 본문에 서명을 넣지 않는다 — Outlook 서명이 자동으로 붙는다
CLOSING = None      # Outlook 서명 첫 줄이 "감사합니다. / 이재현 드림." 이라 본문에 넣으면 중복


# ──────────────────────────────────────────────
# 본문 블록 헬퍼 — 전부 인라인 스타일 HTML 조각을 반환
# ──────────────────────────────────────────────
def p(text):
    """일반 문단."""
    return f"<p style='margin:0 0 12px;'>{text}</p>"


def section(title):
    """번호 소제목: '1. 주요 내용' 형태 그대로 넘긴다."""
    return f"<p style='margin:16px 0 6px;'><b>{title}</b></p>"


def bullets(items):
    """불릿 목록."""
    li = "".join(f"<li style='margin:2px 0;'>{i}</li>" for i in items)
    return f"<ul style='margin:0 0 12px; padding-left:22px;'>{li}</ul>"


def table(headers, rows):
    """얇은 테두리 요약 표. rows는 문자열 2차원 리스트."""
    th = "".join(f"<th style='{_TD} background:#F2F2F2; font-weight:bold; text-align:left;'>{h}</th>"
                 for h in headers)
    trs = ""
    for r in rows:
        tds = "".join(f"<td style='{_TD}'>{c}</td>" for c in r)
        trs += f"<tr>{tds}</tr>"
    return (f"<table style='border-collapse:collapse; {FONT} font-size:{FONT_SIZE};"
            f" color:{TEXT_COLOR}; margin:4px 0 14px;'>"
            f"<tr>{th}</tr>{trs}</table>")


def assemble_body(greeting, blocks, closing=CLOSING, sender=SENDER, attachments_note=None):
    """표준 구조로 본문 완성: 인사 → 내용 블록들 → (첨부 표기) → (맺음·서명).

    맺음·서명(closing/sender)은 기본값 None — Outlook 서명이 자동으로 붙기 때문에
    본문에 또 쓰면 "감사합니다 / 이재현 드림"이 두 번 나온다. 서명이 없는 상대에게
    보내는 등 본문에 직접 넣어야 할 때만 문자열을 넘긴다.

    바깥 구조는 완전한 HTML 문서 + 1칸 표 래퍼다. Outlook은 초안(.eml)을 열 때 서명을
    "본문 첫 블록 다음"에 끼워 넣는데, 본문 전체가 표 한 칸이면 그 자리가 표 바깥(= 본문 끝)이
    되어 인사말과 본문 사이가 갈라지지 않는다. 마지막 빈 문단은 서명이 안착할 자리다.
    """
    parts = [p(greeting)] + list(blocks)
    if attachments_note:
        parts.append(p(f"<b>첨부</b>: {attachments_note}"))
    if closing:
        parts.append(p(closing))
    if sender:
        parts.append(f"<p style='margin:0;'>{sender} 드림</p>")
    inner = "\n".join(parts)
    return (
        "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>"
        f"<body style=\"{BODY_STYLE}\">"
        "<table role='presentation' cellpadding='0' cellspacing='0' border='0'"
        " style='border-collapse:collapse; width:100%;'>"
        f"<tr><td style=\"{BODY_STYLE} padding:0;\">{inner}</td></tr></table>"
        "<p style='margin:0;'>&nbsp;</p>"
        "</body></html>"
    )


# ──────────────────────────────────────────────
# .eml 생성
# ──────────────────────────────────────────────
def _plain_fallback(html):
    """HTML을 못 보는 클라이언트용 텍스트 폴백(태그 제거)."""
    body = re.search(r"<body[^>]*>(.*)</body>", html, re.S | re.I)
    html = body.group(1) if body else html
    txt = re.sub(r"<(br|/p|/tr|/li)[^>]*>", "\n", html)
    txt = re.sub(r"<[^>]+>", "", txt)
    txt = txt.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\n{3,}", "\n\n", txt).strip()


def build_email(subject, body_html, out_path, to="", cc="",
                attachments=None, plain_text=None):
    """.eml Outlook 초안 파일 생성. 반환: 저장 경로."""
    msg = EmailMessage()
    msg["Subject"] = subject
    if to:
        msg["To"] = to
    if cc:
        msg["Cc"] = cc
    # Outlook이 '보내지 않은 새 메일(작성창)'로 열게 하는 핵심 헤더
    msg["X-Unsent"] = "1"
    msg["Date"] = formatdate(localtime=True)   # 없으면 Outlook이 파일을 못 여는 경우가 있다

    msg.set_content(plain_text or _plain_fallback(body_html))
    msg.add_alternative(body_html, subtype="html")

    for path in attachments or []:
        if not os.path.exists(path):
            raise FileNotFoundError(f"첨부 파일 없음: {path}")
        ctype, _ = mimetypes.guess_type(path)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype,
                               filename=os.path.basename(path))

    with open(out_path, "wb") as f:
        BytesGenerator(f).flatten(msg)
    return out_path
