# -*- coding: utf-8 -*-
"""
word_themes.py — 워드 전용 테마·문서유형 레지스트리.

엑셀(excel_theme.PALETTES)과 색을 맞출 필요가 없으므로 워드 테마는 여기서 독립 정의한다.
word_theme.py 가 이 dict 를 읽어 get_theme(name)/DEFAULT_THEME 으로 노출한다.

  • default — 기본 테마. 모노톤(라이트 그레이), 라이트 헤더, 삼선표, 에어리한 보고서 톤.
  • procpa  — 브랜드 블루. default 와 같은 레이아웃에 procpa.co.kr 팔레트
              (딥네이비 #0B1C4A + 프라이머리 블루 #2563EB), 폰트 Pretendard.
  • (별칭) workpaper → default — 구 조서용 테마는 default 로 통합.
              조서·내부 산출물은 new_doc(doc_type="workpaper") 로 구분한다.

표 옵션 키(word_theme.add_table 가 해석):
  table_sides(False=좌우 외곽선 없음) · table_edge_sz(상/하 외곽 굵기 1/8pt) ·
  table_header_sep(헤더 하단/합계 상단 옅은 구분선) · table_sep_sz/table_total_sz(그 구분선 굵기 1/8pt) ·
  table_header_fill/table_header_text(라이트 헤더) ·
  zebra(False=줄무늬 없음)

제목부 스타일(add_title style=)과 문서유형 프리셋(new_doc doc_type=)은
아래 TITLE_STYLES / DOC_TYPES 참조.
"""

_FONT = "맑은 고딕"
_FALLBACK = "Malgun Gothic"
_SIZES = {"title": 24, "subtitle": 11, "h1": 14, "h2": 12,
          "h3": 12, "header": 11, "body": 11, "small": 9}
# 표 공통(삼선표 + 라이트 헤더)
_TABLE = {"zebra": False, "table_sides": False, "table_edge_sz": 4,
          "table_header_sep": True, "table_sep_sz": 4, "table_total_sz": 4}
# 2026-09 사용자 확정: 삼선표 구조는 유지하되 모든 테두리를 일반 굵기(0.5pt=4/8pt)로 통일.
# (종전 상하 2pt·헤더구분 1pt·합계 1.5pt 의 3단계 굵기는 '만든 티'가 나서 폐기)


def _theme(label, *, head, head_sep, band, text, note, accent, negative,
           border_in, edge, hdr_fill, hdr_text, font=_FONT, fallback=_FALLBACK):
    """공통 레이아웃 + 색만 받아 워드 테마 dict 구성."""
    t = {
        "label": label,
        "font_name": font, "font_excel": font, "font_fallback": fallback,
        "sizes": dict(_SIZES),
        "colors": {
            "primary": head, "primary_dark": head, "secondary": head,  # 제목단계 동일색
            "secondary_lt": head_sep, "band": band,
            "subheader": hdr_fill, "surface": "FFFFFF",
            "text": text, "note": note, "accent": accent,
            "negative": negative, "header_font": "FFFFFF",
            "title_color": head, "border_in": border_in, "border_out": edge,
            "table_header_fill": hdr_fill, "table_header_text": hdr_text,
        },
    }
    t.update(_TABLE)
    return t


WORD_THEMES = {
    # 기본 — 모노톤 그레이 (라이트·에어리, 보고서·제안서·조서)
    "default": _theme(
        "기본 — 모노톤 그레이 (보고서·메모·조서)",
        head="333333", head_sep="BFBFBF", band="F4F4F4",
        text="222222", note="F2F2F2", accent="595959", negative="C0392B",
        border_in="DDDDDD", edge="595959", hdr_fill="ECECEC", hdr_text="333333"),

    # 브랜드 블루 — procpa.co.kr 팔레트 (같은 레이아웃, 색·폰트만 교체)
    "procpa": _theme(
        "procpa — 브랜드 블루 (보고서·제안서)",
        head="0B1C4A", head_sep="A8C6FF", band="EFF5FF",
        text="1A2330", note="EFF5FF", accent="2563EB", negative="C0392B",
        border_in="D6E0EE", edge="0B1C4A", hdr_fill="EFF5FF", hdr_text="0B1C4A",
        # Pretendard Variable — static Bold 부재로 인한 가짜 볼드 방지 (qa 린터 검출)
        font="Pretendard Variable", fallback="맑은 고딕"),
}

# 구 이름 → 신 이름 별칭 (opinion-letter 등 하위호환)
WORD_ALIASES = {"workpaper": "default"}

DEFAULT_WORD_THEME = "default"

# ──────────────────────────────────────────────────────────────
# 제목부 스타일 5종 — add_title(style=...) 이 해석 (구현: word_theme.py)
#   bar    제목 + 하단 강조 바 (기본 — 일반 보고서. 두 줄 제목은 고정 줄높이로 압축)
#   side   좌측 세로 악센트 바 + 제목/부제                     — 간단한 메모·검토
#   band   전폭 색 밴드(짙은 배경 + 흰 제목)                   — 제안서·트렌디 산출물
#   center 중앙 정렬 + 상하 가는 선(88% 폭)                    — 공식 의견서·공문·회신문
#   meta   제목(좌) + 문서정보 표(우) + 하단 굵은선            — (구) 조서
#   block  정보표 2×4(회사명│결산일 / 조서번호│작성자, 라벨 음영 격자) 위 + 제목 + 0.5pt 하단선, 폭=본문 폭 — 조서 기본(2026-09-07 최종, 4항목)
# ──────────────────────────────────────────────────────────────
TITLE_STYLES = ("bar", "side", "band", "center", "meta", "block")

# 문서유형 프리셋 — new_doc(doc_type=...) 지정 시 add_title 의 기본 스타일·캡션이 결정된다.
#   caption 은 band 스타일의 문서유형 라벨 기본값 (meta={"caption": ...} 로 덮어씀).
DOC_TYPES = {
    "report":    {"title_style": "bar",    "caption": None},
    "memo":      {"title_style": "side",   "caption": None},
    "proposal":  {"title_style": "band",   "caption": "PROPOSAL"},
    "opinion":   {"title_style": "center", "caption": None},
    "workpaper": {"title_style": "block",  "caption": None},   # 2026-09: meta → block 확정
}
