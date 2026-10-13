# -*- coding: utf-8 -*-
"""
엑셀 테마 — 색/폰트/표 규칙 + openpyxl 헬퍼를 **이 한 파일**에 모았다.
(office-palette 분리 없이 단독 완결. 워드/PPT 스킬도 이 파일의 팔레트를 참조한다.)

■ 색/폰트/표 규칙을 바꾸려면 아래 "편집 영역"의 PALETTES / NUMBER_FORMATS / DEFAULT_THEME 만 고친다. ■
그 아래 매핑·헬퍼 로직은 보통 손댈 필요가 없다.

    from excel_theme import apply_theme
    apply_theme(ws, theme="default", header_row=4, data_range="B4:F9",
                title_cell="B2", number_format_cols={"C": "accounting"})
"""

import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.utils.cell import coordinate_from_string

# ══════════════════════════════════════════════════════════════════════
# ■ 최신 Office 테마 강제 — openpyxl 기본 "Office 2007-2010" 제거 ■
# ══════════════════════════════════════════════════════════════════════
# openpyxl 은 저장 시 항상 구버전(Office 2007-2010) 테마 XML 을 끼워 넣어,
# Excel 의 [페이지 레이아웃]>[테마] 가 "Office 2007-2010" 으로 잡힌다.
# 아래에서 최신(2013+) Office 테마(Calibri Light/Calibri · 모던 색상)를 강제한다.
#   - 이 모듈을 import 하면 openpyxl 기본 저장 경로가 자동으로 최신 테마를 쓴다.
#   - 특정 워크북만 개별 지정하려면 modernize_theme(wb) 를 save() 직전에 호출.
MODERN_OFFICE_THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements><a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
    '</a:clrScheme><a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="맑은 고딕"/>'
    '<a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="맑은 고딕"/>'
    '<a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" rotWithShape="0">'
    '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>'
    '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
)


def modernize_theme(wb):
    """워크북에 최신 Office 테마를 박아 넣는다(save 직전 호출). 반환: wb."""
    wb.loaded_theme = MODERN_OFFICE_THEME
    return wb


# import 시점에 openpyxl 기본 저장 경로(loaded_theme 미설정 시)를 최신 테마로 교체.
# writer.excel 이 `from .theme import theme_xml` 로 값을 복사해 가므로 그 네임스페이스를 직접 패치.
try:
    import openpyxl.writer.excel as _ox_writer
    _ox_writer.theme_xml = MODERN_OFFICE_THEME
except Exception:
    pass

# ══════════════════════════════════════════════════════════════════════
# ■■■ 편집 영역 — 색 / 폰트 / 표 규칙은 여기서 모두 수정한다 ■■■
# ══════════════════════════════════════════════════════════════════════
#
# 의미 토큰(색): primary 헤더/제목·테두리 / primary_dark 주요항목·합계 글자 /
#   secondary·secondary_lt 보조 / band 합계·강조 배경 / subheader 섹션배경 /
#   surface 옅은 면 / text 본문글자 / note 강조박스 / accent 포인트 /
#   negative 음수 / header_font 헤더글자 / title_color 제목글자 / border_in·border_out 테두리
# 엑셀 폰트는 font_excel(없으면 font_name). 엑셀 렌더 플래그:
#   excel_table_style: "rules"(흰배경 헤더+가로구분선) | "grid"(전체 격자)
#   excel_zebra / excel_highlight_neg / excel_neg_red (음수 [Red] 사용)

# 미지정 시 기본 테마. 공식 카탈로그 3종 — 목적에 따라 theme= 로 골라 쓴다:
#   default = 정산표·결산(회색·기본) / procpa = 브랜드 블루(procpa.co.kr) / dcf-valuation = DCF·평가
#   navy·charcoal 은 legacy(카탈로그 제외, 하위호환용 — 기존 감사조서 룩 유지).
#   구 이름은 ALIASES 로 계속 동작: closing→default, valuation→dcf-valuation,
#   audit_charcoal→charcoal, audit_navy→navy (감사조서 빌더 frame.py 하위호환)
# 기본 테마. 환경변수 EXCEL_THEME_DEFAULT 로 프로젝트별 덮어쓰기 가능(예: 감사조서 공간=audit).
DEFAULT_THEME = os.environ.get("EXCEL_THEME_DEFAULT", "default")

PALETTES = {
    # ── DCF·평가용 : Ocean Blue Clean (구 valuation) ─────────────────────
    "dcf-valuation": {
        "label": "DCF·평가 (Ocean Blue Clean)",
        "font_name": "Pretendard",       # 워드/PPT 기본 폰트
        "font_excel": "맑은 고딕",         # 엑셀 폰트
        "font_fallback": "맑은 고딕",
        "sizes": {"title": 20, "subtitle": 13, "h1": 16, "h2": 13,
                  "h3": 11, "header": 11, "body": 11, "small": 9},
        "excel_zebra": False, "excel_highlight_neg": False, "excel_neg_red": False,
        "excel_table_style": "rules",
        "excel_title_size": 12, "excel_size_header": 11, "excel_size_body": 11,
        "excel_row_header": 22, "excel_row_body": 18,
        # 탭 색: 팔레트의 진한→연한 오션블루 4단 + 무채색 원본
        "tabs": {"guide": None, "output": "004889", "calc": "2B579A",
                 "input": "5B9BD5", "pbc": "D6E4F0", "raw": "D9D9D9"},
        "colors": {
            "primary": "004889", "primary_dark": "1F3864", "secondary": "2B579A",
            "secondary_lt": "5B9BD5", "band": "F5F8FC", "subheader": "F5F8FC",
            "surface": "F5F8FC", "text": "3D4F5F", "note": "F5F8FC",
            "accent": "5B9BD5", "negative": "5A6A7A", "header_font": "FFFFFF",
            "title_color": "1F3864", "border_in": "D6E4F0", "border_out": "004889",
        },
    },
    # ── 정산표·결산·감사조서 : 고객사 정산표 (구 settlement) ────────
    #   statement 스타일: 회색(F2F2F2) bold 헤더 + 표 상/하 medium 테두리 + 격자 없음.
    #   하드코딩 셀 #FFFBEF, 음수 빨강 괄호, 맑은 고딕. (FY26.1Q 파일 분석)
    "default": {
        "label": "default (정산표·결산 — 회색 statement)",
        "font_name": "맑은 고딕", "font_excel": "맑은 고딕", "font_fallback": "Malgun Gothic",
        "sizes": {"title": 12, "subtitle": 11, "h1": 14, "h2": 12,
                  "h3": 11, "header": 11, "body": 11, "small": 10},
        "excel_zebra": False, "excel_highlight_neg": False, "excel_neg_red": True,
        "excel_table_style": "frame", "excel_title_fill": True,
        "excel_currency_format": "accounting",    # 통화 기본 = 원 단위 회계서식 (사용자 확정 2026-09-14, 백만원은 number_cols 로 지정)
        "excel_title_size": 11, "excel_size_header": 11, "excel_size_body": 11,
        "excel_row_header": 20, "excel_row_body": 16,
        "colors": {
            "primary": "404040",       # (워드/PPT 헤딩·헤더·타이틀) 진회색
            "primary_dark": "262626",  # 강조/소계·합계 글자
            "secondary": "808080", "secondary_lt": "BFBFBF",
            "band": "F2F2F2",          # 줄무늬 배경(현재 미사용)
            "subheader": "FFFFFF",     # 섹션행 배경 없음(굵은 글씨로만 구분)
            "surface": "FFFFFF",
            "text": "000000",          # 본문 검정
            "note": "FFFBEF",          # 하드코딩(입력) 셀 — 크림
            "linked": "F2F2F2",        # 참조값(타시트 연결) 셀 — 연회색
            "accent": "C00000",        # 포인트(진한 빨강)
            "negative": "FF0000",      # 음수 빨강
            "header_font": "FFFFFF",   # (워드/PPT 헤더 글자) 흰색
            "title_color": "000000",
            "border_in": "BFBFBF",     # (쓰면) 얇은 선
            "border_out": "000000",    # 엑셀 표 상/하 굵은(medium) 선
            # ── 엑셀 전용 오버라이드 (정산표 statement 룩) ──
            "excel_header_fill": "F2F2F2",  # 본문 기본 헤더(회색)
            "excel_header_text": "000000",  # 헤더 글자(검정 bold)
            "excel_header2_fill": "DDE3E8",  # 본문 세컨더리 헤더(연청회색)
            "excel_header2_text": "000000",
            "excel_total_fill": "F2F2F2",    # 소계/합계 배경 = 기본 헤더색
            "title_bg": "393939",           # 제목(B2) 헤더 바(진회색, 데이터 끝열까지)
            "title_font_color": "FFFFFF",
        },
    },

    # ── 감사조서 : 색·폰트·표 규칙 모두 default 와 동일 (사용자 확정 2026-09-08) ──
    #   default 와 다른 것은 "쓰는 법" 두 가지뿐: ① 시트 제목 밴드를 두지 않고 상단에 조서 헤더 6항목
    #   (write_audit_header: 회사명·조서번호·시트명 / 결산일·조서명·작성자, 병합 없음, 색 없음)을 둔다
    #   ② 섹션 소제목("1. 총괄표"…)을 메인헤더 밴드(title_bg #393939 + 흰 굵은 글씨, write_section_bar)로 쓴다.
    "audit": {
        "label": "audit (감사조서 — default 색, 소제목=메인헤더 밴드)",
        "font_name": "맑은 고딕", "font_excel": "맑은 고딕", "font_fallback": "Malgun Gothic",
        "sizes": {"title": 12, "subtitle": 11, "h1": 14, "h2": 12,
                  "h3": 11, "header": 11, "body": 11, "small": 10},
        "excel_zebra": False, "excel_highlight_neg": False, "excel_neg_red": True,
        "excel_table_style": "frame", "excel_title_fill": True,
        "excel_currency_format": "accounting",
        "excel_title_size": 11, "excel_size_header": 11, "excel_size_body": 11,
        "excel_row_header": 20, "excel_row_body": 16,
        "colors": {
            "primary": "404040", "primary_dark": "262626",
            "secondary": "808080", "secondary_lt": "BFBFBF",
            "band": "F2F2F2", "subheader": "FFFFFF", "surface": "FFFFFF", "text": "000000",
            "note": "FFFBEF", "linked": "F2F2F2",
            "accent": "FF0000",        # 틱마크(R·GL·F·PR·Ref.) 빨강 글씨 — default 의 C00000 대신 표준조서 빨강
            "negative": "FF0000", "header_font": "FFFFFF", "title_color": "000000",
            "border_in": "BFBFBF", "border_out": "000000",
            "excel_header_fill": "F2F2F2", "excel_header_text": "000000",
            "excel_header2_fill": "DDE3E8", "excel_header2_text": "000000",
            "excel_total_fill": "F2F2F2",
            "title_bg": "393939", "title_font_color": "FFFFFF",   # 제목 밴드 = 시트 소제목 밴드(메인헤더)
        },
    },

    # ── PROCPA 브랜드 : default 의 statement 레이아웃 + procpa.co.kr 블루 팔레트 ──
    #   딥네이비(#0B1C4A) 타이틀 바 + 블루 틴트(#EFF5FF) 헤더 + 프라이머리 블루(#2563EB) 악센트.
    #   입력(note)·연결(linked)·음수(negative)는 기능색이라 default 와 동일하게 유지.
    "procpa": {
        "label": "procpa (브랜드 블루 — 딥네이비 & 블루)",
        "font_name": "Pretendard", "font_excel": "맑은 고딕", "font_fallback": "맑은 고딕",
        "sizes": {"title": 12, "subtitle": 11, "h1": 14, "h2": 12,
                  "h3": 11, "header": 11, "body": 11, "small": 10},
        "excel_zebra": False, "excel_highlight_neg": False, "excel_neg_red": True,
        "excel_table_style": "frame", "excel_title_fill": True,
        "excel_currency_format": "accounting",
        "excel_title_size": 12, "excel_size_header": 11, "excel_size_body": 11,
        "excel_row_header": 20, "excel_row_body": 16,
        # 탭 색: 팔레트의 진한→연한 브랜드색 4단 + 무채색 원본
        "tabs": {"guide": None, "output": "0B1C4A", "calc": "2563EB",
                 "input": "94A3B8", "pbc": "C9D4E5", "raw": "D9D9D9"},
        "colors": {
            "primary": "2563EB",       # (워드/PPT 헤딩·헤더) 프라이머리 블루
            "primary_dark": "0B1C4A",  # 강조/소계·합계 글자 — 딥네이비
            "secondary": "4A5160", "secondary_lt": "94A3B8",
            "band": "EFF5FF",          # 줄무늬/합계 배경 — 블루 틴트
            "subheader": "FFFFFF",     # 섹션행 배경 없음(굵은 글씨로만 구분)
            "surface": "FFFFFF",
            "text": "000000",
            "note": "FFFBEF",          # 하드코딩(입력) 셀 — 크림 (기능색, 전 테마 공통)
            "linked": "F2F2F2",        # 참조값(타시트 연결) 셀 — 연회색 (기능색)
            "accent": "2563EB",        # 포인트
            "negative": "FF0000",      # 음수 빨강 (회계 관행)
            "header_font": "FFFFFF",
            "title_color": "0B1C4A",
            "border_in": "C9D4E5",     # 블루그레이 헤어라인
            "border_out": "0B1C4A",    # 표 상/하 medium 선 — 딥네이비
            # ── 엑셀 전용 오버라이드 (statement 룩) ──
            "excel_header_fill": "EFF5FF",   # 본문 기본 헤더(블루 틴트)
            "excel_header_text": "0B1C4A",
            "excel_header2_fill": "F3F5F8",  # 본문 세컨더리 헤더(그레이 서피스)
            "excel_header2_text": "0B1C4A",
            "excel_total_fill": "EFF5FF",    # 소계/합계 배경 = 기본 헤더색
            "title_bg": "0B1C4A",            # 제목(B2) 헤더 바(딥네이비)
            "title_font_color": "FFFFFF",
        },
    },

    # ── [legacy] 감사조서용 팔레트 2종 — 카탈로그 제외, 하위호환용(기존 조서 룩 유지) ──
    #   공통: 맑은 고딕, 입력=크림(note)·연결=옅은 동색(linked), grid 테두리(외곽 medium=진한색).
    #   고도화(2026-06): 섹션·헤더 계층 대비 정교화 + 테마별 linked 색 추가.
    # 표=frame(삼선표: 좌우 외곽선 없음·상하 굵게·내부 얇게) · 헤더 3색(제목/본문/세컨더리)
    # 셀배경: 하드코딩=note(크림) / 수식=흰배경(없음) / 참조값=linked(#F2F2F2)
    "navy": {
        "label": "navy (네이비 & 골드) — legacy",
        "legacy": True,
        "font_name": "맑은 고딕", "font_excel": "맑은 고딕", "font_fallback": "Malgun Gothic",
        "sizes": {"title": 12, "subtitle": 11, "h1": 14, "h2": 12, "h3": 11, "header": 11, "body": 11, "small": 9},
        # excel_highlight_neg=False → 음수는 숫자서식 [Red](기본 빨강)만, bold/별색 오버레이 없음
        "excel_zebra": False, "excel_highlight_neg": False, "excel_neg_red": True,
        "excel_table_style": "frame", "excel_title_fill": True,
        "excel_title_size": 12, "excel_size_header": 11, "excel_size_body": 11,
        "excel_row_header": 22, "excel_row_body": 20,
        "colors": {
            "primary": "1F3A5F", "primary_dark": "16293F", "secondary": "3A5172", "secondary_lt": "8AA0BC",
            "band": "EEF2F7", "subheader": "FFFFFF", "surface": "FFFFFF", "text": "1A2330",
            "note": "FFFBEF", "linked": "F2F2F2", "accent": "C9A227", "negative": "FF0000", "header_font": "FFFFFF",
            "title_color": "1F3A5F", "border_in": "C5CEDA", "border_out": "1F3A5F",
            "excel_header_fill": "E1E8F1", "excel_header_text": "16293F",      # 본문 기본 헤더
            "excel_header2_fill": "DCE6F4", "excel_header2_text": "16293F",    # 본문 세컨더리 헤더
            "excel_total_fill": "FFFFFF",                                      # 소계/합계 배경 없음(윗선+굵게로만)
            "title_bg": "1F3A5F", "title_font_color": "FFFFFF",               # 제목(B2) 헤더
            # subheader=FFFFFF → 섹션행 별도 배경색 없음(굵은 글씨로만 구분)
        },
    },
    "charcoal": {
        "label": "charcoal (차콜 & 틸) — legacy",
        "legacy": True,
        "font_name": "맑은 고딕", "font_excel": "맑은 고딕", "font_fallback": "Malgun Gothic",
        "sizes": {"title": 12, "subtitle": 11, "h1": 14, "h2": 12, "h3": 11, "header": 11, "body": 11, "small": 9},
        # excel_highlight_neg=False → 음수는 굵게/별색 오버레이 없이 숫자서식 [Red](기본 빨강)만
        "excel_zebra": False, "excel_highlight_neg": False, "excel_neg_red": True,
        "excel_table_style": "frame", "excel_title_fill": True,
        "excel_title_size": 12, "excel_size_header": 11, "excel_size_body": 11,
        "excel_row_header": 22, "excel_row_body": 20,
        "colors": {
            "primary": "2A2E35", "primary_dark": "1A1D22", "secondary": "4A4F58", "secondary_lt": "9AA0AA",
            "band": "EDEFF2", "subheader": "FFFFFF", "surface": "FFFFFF", "text": "22262B",
            "note": "FFFBEF", "linked": "F2F2F2", "accent": "2E8B7F", "negative": "FF0000", "header_font": "FFFFFF",
            "title_color": "2A2E35", "border_in": "D2D6DC", "border_out": "2A2E35",
            "excel_header_fill": "D6DEE6", "excel_header_text": "2A2E35",      # 본문 기본 헤더(슬레이트)
            "excel_header2_fill": "E8ECEF", "excel_header2_text": "2A2E35",    # 본문 세컨더리(연그레이)
            "excel_total_fill": "FFFFFF",                                      # 소계/합계 배경 없음(윗선+굵게로만)
            "title_bg": "2A2E35", "title_font_color": "FFFFFF",               # 제목(B2) 헤더
            # subheader=FFFFFF → 섹션행 별도 배경색 없음(굵은 글씨로만 구분)
        },
    },
}

# 구 이름 → 신 이름 별칭 (감사조서 빌더 frame.py·valuation-tools 등 하위호환)
ALIASES = {"valuation": "dcf-valuation",
           "감사조서": "audit", "workpaper": "audit", "standard": "audit",
           "audit_charcoal": "charcoal", "audit_navy": "navy", "closing": "default",
           "default - charcoal": "charcoal", "default - navy": "navy"}

# 공통 숫자 서식 (음수=빨강 괄호, 0=대시). 테마가 [Red] 제거 옵션을 가질 수 있음.
NUMBER_FORMATS = {
    "accounting":   "#,###,##0;[Red](#,###,##0);-",
    "percent_acct": '0.00%_);[Red](0.00%);"-"_);@_)',
    "thousands":    "#,##0;[Red](#,##0)",
    "percent":      "0.0%;[Red]-0.0%",
    "currency_won": '#,##0"원";[Red](#,##0"원")',
    "million_won":  '#,##0,,"백만원";[Red](#,##0,,"백만원")',
    "decimal":      "#,##0.00;[Red](#,##0.00)",
    "date":         "yyyy-mm-dd",
    # 추가: 모델링/배수/스케일/증감
    "multiple":     "0.0x",                                  # 배수 (8.5x)
    "million":      "#,##0,,;[Red](#,##0,,);-",              # 백만 스케일(라벨 없음)
    "billion":      "#,##0,,,;[Red](#,##0,,,);-",            # 십억 스케일
    "change":       "+#,##0;[Red]-#,##0;-",                  # 증감(부호)
    "bp":           '#,##0" bp"',                            # 베이시스포인트
}

# ── 시트 탭 색 (사용자 확정 2026-09-10) ────────────────────────────────
#   두 축: **진하기 = 자리** — 진할수록 파일의 중심(결과), 연할수록 바깥(원천 자료).
#          **무채색 = 우리 산출물이 아님** — 원본(raw)은 어느 테마에서든 회색으로 고정.
#   색은 그 테마가 이미 가진 토큰에서만 고른다. 탭을 위해 새 색을 만들지 않는다.
#   · 회색 테마(default·audit·legacy) = Office 표준 회색 램프가 그대로 위계가 된다.
#   · 색 있는 테마(procpa·dcf-valuation) = 팔레트의 진한→연한 브랜드색 4단 + 무채색 원본
#     (팔레트 안 "tabs" 로 오버라이드. 없으면 아래 회색 램프를 쓴다).
#   안내·표지는 색을 주지 않는다 — 색 없는 탭이 하나 있어야 나머지 색이 의미를 갖는다.
TAB_COLORS = {
    "guide":  None,        # 안내·표지 — 색 없음
    "output": "393939",    # 결과·총괄 (= default title_bg, 제목 밴드와 같은 색)
    "calc":   "808080",    # 계산·집계
    "input":  "A6A6A6",    # 입력 — 작성자(파라미터·판단·전기 이월)
    "pbc":    "BFBFBF",    # 입력 — 회사 제공(PBC)
    "raw":    "D9D9D9",    # 원본·참고(열람만, 편집 금지)
}
TAB_ROLE_ALIASES = {       # 감사(pbc)·평가(input_ext) 어느 용어로 불러도 받는다
    "안내": "guide", "표지": "guide", "cover": "guide",
    "결과": "output", "총괄": "output", "result": "output",
    "계산": "calc", "집계": "calc",
    "입력": "input", "작성자": "input", "input_own": "input", "assumption": "input",
    "회사": "pbc", "외부": "pbc", "input_ext": "pbc", "client": "pbc",
    "원본": "raw", "참고": "raw", "source": "raw",
}
# ══════════════════════════════════════════════════════════════════════
# ■■■ 편집 영역 끝 — 아래는 로직 (보통 수정 불필요) ■■■
# ══════════════════════════════════════════════════════════════════════

__all__ = [
    "apply_theme", "style_header_row", "apply_borders",
    "set_number_formats", "zebra_stripes", "freeze_below_header",
    "autofit_columns", "highlight_negatives", "highlight_threshold",
    "style_total_row", "style_subheader_row", "highlight_hardcoded",
    "write_table", "mark_cells", "write_caption", "sheet_label", "title_only", "fit_rows", "write_header", "write_section_bar", "write_tickmark", "write_audit_header", "AUDIT_HEADER_FIELDS", "HEADER_STYLES", "write_cover_sheet", "write_title", "COVER_FIELDS", "COVER_SHEET",
    "set_tab", "TAB_COLORS", "TAB_ROLE_ALIASES",
    "PALETTES", "NUMBER_FORMATS", "THEMES", "DEFAULT_THEME",
    "get_palette", "get_theme", "color", "theme_from_palette", "MAX_COL_WIDTH",
    "MODERN_OFFICE_THEME", "modernize_theme",
]


# ── 팔레트 접근 (워드/PPT 도 사용: 의미 토큰 dict 그대로) ─────────────
def _canon(name):
    """구 이름(ALIASES)을 신 이름으로 정규화."""
    return ALIASES.get(name, name)


def get_palette(name=None):
    """팔레트(의미 토큰) dict 반환. None/미존재 시 기본. (구 이름 별칭 허용)"""
    if not name:
        return PALETTES[DEFAULT_THEME]
    name = _canon(name)
    if name not in PALETTES:
        raise KeyError(f"알 수 없는 테마 '{name}'. 사용 가능: {', '.join(PALETTES)}")
    return PALETTES[name]


def color(name, role):
    """팔레트 name 의 의미색 role 을 hex 로."""
    return get_palette(name)["colors"][role]


# ── 의미 팔레트 → 엑셀 토큰(header_bg 등) 매핑 ───────────────────────
def _excel_theme(palette_name):
    return theme_from_palette(PALETTES[palette_name])


def theme_from_palette(p):
    """팔레트 dict → 엑셀 토큰 dict (샘플·변형 테마용: PALETTES 에 없는 팔레트도 바로 쓸 수 있다)."""
    c, s = p["colors"], p["sizes"]
    nf = dict(NUMBER_FORMATS)
    if not p.get("excel_neg_red", True):           # 음수 [Red] 제거(괄호만)
        nf = {k: v.replace("[Red]", "") for k, v in nf.items()}
    style = p.get("excel_table_style", "grid")
    rules_like = style in ("rules", "statement")     # 격자 없음·합계 가로선 방식
    return {
        "font_name":          p.get("font_excel", p["font_name"]),
        "font_size_title":    p.get("excel_title_size", 11),   # 조서 제목 밴드 11pt (사용자 확정 2026-09-07)
        "font_size_header":   p.get("excel_size_header", s["header"]),
        "font_size_body":     p.get("excel_size_body", s["body"]),
        "header_bg":          c.get("excel_header_fill", c["primary"]),
        "header_font_color":  c.get("excel_header_text", c["header_font"]),
        "header2_bg":         c.get("excel_header2_fill", c["subheader"]),    # 본문 세컨더리 헤더
        "header2_font":       c.get("excel_header2_text", c["text"]),
        "title_color":        c["title_color"],
        "band_fill":          c["band"],
        "subheader_fill":     c["subheader"],
        "total_fill":         c.get("excel_total_fill", c["band"]),   # 소계/합계 배경(기본=band, 테마 지정 시 헤더색 등)
        "note_fill":          c["note"],
        "linked_fill":        c.get("linked", "CCECFF"),     # 타시트 연결 셀(파랑)
        "todo_fill":          c.get("todo", "FFFF00"),       # 미입수 - 아직 회사에서 못 받은 자료(노랑)
        "currency_format":    p.get("excel_currency_format", "accounting"),
        "accent_color":       c["negative"],
        "border_color":       c["border_in"],
        "border_color_outer": c["border_out"],
        "border_outer_style": p.get("excel_border_outer_style", "medium"),  # grid/frame 외곽선 굵기
        "total_style":        p.get("excel_total_style", "default"),        # 합계행: default | double
        "row_height_header":  p.get("excel_row_header", 22),
        "row_height_body":    p.get("excel_row_body", 18),
        "number_formats":     nf,
        "excel_zebra":         p.get("excel_zebra", True),
        "excel_highlight_neg": p.get("excel_highlight_neg", True),
        "header_fill":         style != "rules",
        "header_text_color":   (c["primary"] if style == "rules"
                                else c.get("excel_header_text", c["header_font"])),
        "border_mode":         style,
        "text_color":          c["text"],
        "title_fill":          p.get("excel_title_fill", style == "rules"),
        "title_bg":            c.get("title_bg", c["primary"]),
        "title_font_color":    c.get("title_font_color", c["header_font"]),
        "total_font_color":    (c["primary_dark"] if rules_like else None),
        "tab_colors":          {**TAB_COLORS, **p.get("tabs", {})},  # 시트 탭 색(역할→hex, guide=None)
        "legacy":              p.get("legacy", False),   # 카탈로그 제외(하위호환) 표시
    }


THEMES = {name: _excel_theme(name) for name in PALETTES}


def get_theme(name=None):
    """엑셀 토큰 dict 반환. None/미존재 시 기본. (구 이름 별칭 허용)"""
    if not name:
        return THEMES[DEFAULT_THEME]
    name = _canon(name)
    if name not in THEMES:
        raise KeyError(f"알 수 없는 테마 '{name}'. 사용 가능: {', '.join(THEMES)}")
    return THEMES[name]


# ──────────────────────────────────────────────────────────────
# 내부 유틸
# ──────────────────────────────────────────────────────────────
def _bounds(ws, data_range):
    """data_range 문자열(또는 None)을 (min_col,min_row,max_col,max_row)로.
    None 이면 ws.dimensions(사용 영역)를 쓴다."""
    if data_range is None:
        data_range = ws.calculate_dimension()
    return range_boundaries(data_range)


def _fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


# ──────────────────────────────────────────────────────────────
# 1. 헤더 스타일
# ──────────────────────────────────────────────────────────────
def style_header_row(ws, theme, header_row, min_col, max_col, secondary=False):
    """헤더 행 스타일.

    header_fill=True  : 배경색 채움 + 흰(또는 지정) 글씨 (격자 스타일, 기본)
    header_fill=False : 흰 배경 + primary 굵은 글씨 + 하단 medium 선만
                        (Ocean Blue '가로 구분선' 스타일)
    secondary=True    : 본문 세컨더리 헤더 색(header2_bg/header2_font) — 한 시트에 표가
                        2차로 나올 때 본문 기본 헤더와 구분.
    """
    mode = theme.get("border_mode", "grid")
    header_fill = theme.get("header_fill", True)
    if secondary:
        bg = theme.get("header2_bg", theme["header_bg"])
        text_color = theme.get("header2_font",
                               theme.get("header_text_color", theme["header_font_color"]))
    else:
        bg = theme["header_bg"]
        text_color = theme.get("header_text_color", theme["header_font_color"])
    oc = theme["border_color_outer"]
    font = Font(name=theme["font_name"], size=theme["font_size_header"],
                bold=True, color=text_color)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = _fill(bg) if header_fill else None
    # 헤더 테두리: rules(흰헤더)=하단 medium / statement=상단 medium + 하단 thin
    top_b = bot_b = None
    if mode == "rules" and not header_fill:
        bot_b = Side(style="medium", color=oc)
    elif mode == "statement":
        top_b = Side(style="medium", color=oc)
        bot_b = Side(style="thin", color=oc)
    for col in range(min_col, max_col + 1):
        c = ws.cell(row=header_row, column=col)
        if fill is not None:
            c.fill = fill
        c.font = font
        c.alignment = align
        if top_b is not None or bot_b is not None:
            c.border = Border(top=top_b, bottom=bot_b)


# ──────────────────────────────────────────────────────────────
# 2. 테두리 · 정렬 · 행높이
# ──────────────────────────────────────────────────────────────
def apply_borders(ws, theme, min_col, min_row, max_col, max_row):
    """표 내부 얇은 선 + 외곽 굵은 선.

    border_mode="rules"     : 격자 없음(헤더밑/소계/합계 가로선만). Ocean Blue.
    border_mode="statement" : 격자 없음 + 표 맨 위/맨 아래만 medium 굵은선. 정산표.
    border_mode="frame"     : 삼선표 — 좌/우 외곽선 없음, 상/하 굵게(medium), 내부 얇게(thin).
    기본 "grid"             : 외곽 medium + 내부 thin 전체 격자.
    """
    mode = theme.get("border_mode", "grid")
    if mode == "rules":
        return
    if mode == "statement":
        med = Side(style="medium", color=theme["border_color_outer"])
        for col in range(min_col, max_col + 1):
            tc = ws.cell(row=min_row, column=col); pt = tc.border
            tc.border = Border(left=pt.left, right=pt.right, top=med, bottom=pt.bottom)
            bc = ws.cell(row=max_row, column=col); pb = bc.border
            bc.border = Border(left=pb.left, right=pb.right, top=pb.top, bottom=med)
        return
    if mode == "frame":
        thin = Side(style="thin", color=theme["border_color"])
        thick = Side(style="medium", color=theme["border_color_outer"])
        for r in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                left = None if col == min_col else thin       # 좌/우 외곽선 없음
                right = None if col == max_col else thin
                top = thick if r == min_row else thin           # 상/하 굵게
                bottom = thick if r == max_row else thin
                ws.cell(row=r, column=col).border = Border(
                    left=left, right=right, top=top, bottom=bottom)
        return
    thin = Side(style="thin", color=theme["border_color"])
    thick = Side(style=theme.get("border_outer_style", "medium"), color=theme["border_color_outer"])
    for r in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            left = thick if col == min_col else thin
            right = thick if col == max_col else thin
            top = thick if r == min_row else thin
            bottom = thick if r == max_row else thin
            ws.cell(row=r, column=col).border = Border(
                left=left, right=right, top=top, bottom=bottom)


import re

_CRIT_FUNCS = re.compile(r"(?:SUMIFS?|COUNTIFS?|AVERAGEIFS?|MAXIFS|MINIFS)\([^()]*\)", re.I)

def _horiz_for(value):
    """정렬 규칙(사용자 확정 2026-09-04): 숫자·비율·수식은 오른쪽, 텍스트는 왼쪽.
    수식("=…")은 값(숫자)으로 보아 오른쪽 정렬한다. 예외는 *문자열을 반환*하는 수식뿐:
    TEXT(…)·& 연결, 또는 SUMIFS/COUNTIFS 류의 조건 문자열을 걷어낸 뒤에도 빈 문자열("")이
    아닌 문자열 리터럴이 남는 수식(예: =IF(A1>0,"수용","조사"), =IF(…,"OK","확인")).
    =IF(K22="","",ROUND(…)) 처럼 빈 문자열만 있으면 숫자 수식으로 보아 오른쪽."""
    if isinstance(value, bool):
        return "left"
    if isinstance(value, (int, float)):
        return "right"
    if isinstance(value, str) and value.startswith("="):
        up = value.upper()
        if "TEXT(" in up or "&" in value:
            return "left"
        body = value
        for _ in range(5):
            nb = _CRIT_FUNCS.sub("", body)
            if nb == body:
                break
            body = nb
        lits = re.findall(r'"([^"]*)"', body)
        return "right" if all(l == "" for l in lits) else "left"
    return "left"


def style_body(ws, theme, header_row, min_col, min_row, max_col, max_row):
    """본문 폰트 · 세로 가운데정렬 · 행높이. 숫자는 오른쪽, 텍스트는 왼쪽 정렬."""
    body_font = Font(name=theme["font_name"], size=theme["font_size_body"],
                     color=theme.get("text_color"))
    body_start = max(min_row, header_row + 1)
    for r in range(body_start, max_row + 1):
        for col in range(min_col, max_col + 1):
            c = ws.cell(row=r, column=col)
            c.font = body_font
            c.alignment = Alignment(horizontal=_horiz_for(c.value), vertical="center")


# ──────────────────────────────────────────────────────────────
# 3. 줄무늬 (zebra)
# ──────────────────────────────────────────────────────────────
def zebra_stripes(ws, theme, header_row, min_col, min_row, max_col, max_row):
    """헤더 아래 본문에서 짝수번째 데이터 행에 band_fill 적용."""
    fill = _fill(theme["band_fill"])
    body_start = max(min_row, header_row + 1)
    for i, r in enumerate(range(body_start, max_row + 1)):
        if i % 2 == 1:  # 0-based 두번째 행부터 번갈아
            for col in range(min_col, max_col + 1):
                ws.cell(row=r, column=col).fill = fill


# ──────────────────────────────────────────────────────────────
# 4. 틀 고정
# ──────────────────────────────────────────────────────────────
def freeze_below_header(ws, header_row, min_col):
    """헤더 바로 아래 + 첫 데이터 열 기준으로 틀 고정."""
    ws.freeze_panes = ws.cell(row=header_row + 1, column=min_col)


# ──────────────────────────────────────────────────────────────
# 5. 숫자 / 통화 서식
# ──────────────────────────────────────────────────────────────
def set_number_formats(ws, theme, col_formats, min_row, max_row, header_row):
    """열별 숫자 서식 적용.

    col_formats: {"C": "thousands", "D": "percent", 3: "currency_won"}
      값은 themes.number_formats 의 키이거나, 직접 number_format 문자열.
    """
    fmts = theme["number_formats"]
    body_start = max(min_row, header_row + 1)
    for col_key, fmt_name in col_formats.items():
        col = col_key if isinstance(col_key, int) else column_index_from_string(col_key)
        fmt = fmts.get(fmt_name, fmt_name)  # 키가 아니면 원문 그대로
        for r in range(body_start, max_row + 1):
            cell = ws.cell(row=r, column=col)
            cell.number_format = fmt
            # 숫자 서식이 지정된 열은 값·비율·수식 모두 오른쪽 정렬 (텍스트 반환 수식 제외)
            v = cell.value
            if v is None or _horiz_for(v) == "right" or (isinstance(v, str) and v.startswith("=") and "TEXT(" not in v.upper() and "&" not in v):
                cell.alignment = Alignment(horizontal="right", vertical="center")


# ──────────────────────────────────────────────────────────────
# 6. 자동 열너비
# ──────────────────────────────────────────────────────────────
MAX_COL_WIDTH = 30   # 열너비 상한(사용자 확정 2026-09-08): 한 열을 이보다 길게 늘리지 않는다. 긴 문장은 wrap·병합으로.


def autofit_columns(ws, min_col, min_row, max_col, max_row, min_width=8, max_width=MAX_COL_WIDTH):
    """내용 길이 기반 열너비. 한글/전각은 폭 2로 계산."""
    for col in range(min_col, max_col + 1):
        longest = 0
        for r in range(min_row, max_row + 1):
            v = ws.cell(row=r, column=col).value
            if v is None:
                continue
            width = sum(2 if ord(ch) > 0x1100 else 1 for ch in str(v))
            longest = max(longest, width)
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(min_width, min(longest + 2, max_width))


# ──────────────────────────────────────────────────────────────
# 7. 조건부 서식 (음수 / 임계값 강조)
# ──────────────────────────────────────────────────────────────
def highlight_negatives(ws, theme, data_range):
    """data_range 안의 음수를 accent_color 글자색 + 굵게 강조."""
    font = Font(color=theme["accent_color"], bold=True)
    rule = CellIsRule(operator="lessThan", formula=["0"], font=font)
    ws.conditional_formatting.add(data_range, rule)


def highlight_threshold(ws, theme, data_range, operator, threshold):
    """임계값 조건부서식. operator: 'greaterThan','lessThan','equal' 등."""
    fill = _fill(theme["band_fill"])
    rule = CellIsRule(operator=operator, formula=[str(threshold)], fill=fill)
    ws.conditional_formatting.add(data_range, rule)


# ──────────────────────────────────────────────────────────────
# 8. 소계 / 합계 행 · 섹션 헤더 (정산표 관행)
# ──────────────────────────────────────────────────────────────
def style_total_row(ws, theme, row, min_col, max_col, fill=True, secondary=False):
    """소계/합계 행 강조. secondary=True 면 배경을 서브표 헤더색(header2_bg)으로 — 서브표의 헤더·합계 색 일치.
    표 맨 아래(삼선표 하단 medium 선)에 오는 합계행은 하단 medium 을 그대로 둔다(thin 으로 덮지 않음, 2026-09-08).

    기본(grid) : medium 상단선 + thin 하단선 + 굵은 글씨 (+선택 배경).
    rules 모드 : 소계=상단 thin 선만, 합계(fill=True)=상단 thin + 하단 medium + 배경.
                 글자색은 total_font_color(보통 네이비). (Ocean Blue 관행)
    """
    rules = theme.get("border_mode", "grid") in ("rules", "statement")
    oc = theme["border_color_outer"]
    if theme.get("total_style") == "double":      # 표준조서: 윗선 thin + 아랫선 double(합계) / thin(소계)
        top = Side(style="thin", color=oc)
        bottom = Side(style="double", color=oc) if fill else Side(style="thin", color=oc)
    elif rules:
        top = Side(style="thin", color=oc)
        bottom = Side(style="medium", color=oc) if fill else None
    else:
        top = Side(style="medium", color=oc)
        bottom = Side(style="thin", color=oc)
    tf = theme.get("header2_bg", theme["header_bg"]) if secondary else theme.get("total_fill", theme["band_fill"])
    bg = _fill(tf) if fill else None
    fcolor = theme.get("total_font_color")
    for col in range(min_col, max_col + 1):
        c = ws.cell(row=row, column=col)
        c.font = Font(name=theme["font_name"], size=theme["font_size_body"],
                      bold=True, color=fcolor)
        # 기존 좌우 테두리는 유지하고 상/하만 덮어쓴다. 하단이 이미 굵은 선(표 맨 아래)이면 그대로 둔다.
        prev = c.border
        keep_bottom = prev.bottom is not None and prev.bottom.style in ("medium", "thick", "double")
        c.border = Border(left=prev.left, right=prev.right,
                          top=top, bottom=(prev.bottom if keep_bottom or bottom is None else bottom))
        if bg is not None:
            c.fill = bg


def style_subheader_row(ws, theme, row, min_col, max_col):
    """섹션 구분 행: subheader_fill 배경 + 굵은 글씨.
    rules 모드에선 글자색을 total_font_color(네이비)로 맞춘다."""
    bg = _fill(theme.get("subheader_fill", theme["band_fill"]))
    fcolor = theme.get("total_font_color")
    for col in range(min_col, max_col + 1):
        c = ws.cell(row=row, column=col)
        c.fill = bg
        c.font = Font(name=theme["font_name"], size=theme["font_size_body"],
                      bold=True, color=fcolor)


# 문장부호 규칙(사용자 확정 2026-09-11): 엑셀 셀 텍스트에 긴 대시를 쓰지 않는다(가운뎃점 · 은 허용).
#   — – ― (긴/중간 대시) → "-"    ("연령분석 — 잔액" → "연령분석 - 잔액")
PUNCT_DASHES = "—–―"
_DASH_RE = re.compile(r"[" + PUNCT_DASHES + r"]")


def plain_punct(value):
    """셀 텍스트의 긴 대시를 "-" 로 바꾼다. 문자열이 아니거나 수식(=…)이면 그대로."""
    if not isinstance(value, str) or value.startswith("="):
        return value
    return _DASH_RE.sub("-", value)


def normalize_punct(ws, min_row, min_col, max_row, max_col):
    """범위 안의 텍스트 셀에 plain_punct 를 적용한다(수식·숫자는 건드리지 않음). 반환: 바꾼 셀 수."""
    n = 0
    for r in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            c = ws.cell(row=r, column=col)
            v = plain_punct(c.value)
            if v != c.value:
                c.value = v; n += 1
    return n


def write_caption(ws, cell, text, theme=DEFAULT_THEME):
    """표 소제목(캡션) 규칙(사용자 확정 2026-09-04): 본문 폰트(맑은 고딕) 11pt 굵게, 검정.
    캡션 행 바로 아래 1행은 비워 두고 그 다음 행에 표 헤더를 둔다
    (예: B12 캡션 → 13행 공백 → 14행 헤더). 캡션은 표 영역 밖이므로 apply_theme 대상이 아니다."""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    c = ws[cell]
    c.value = plain_punct(text)
    c.font = Font(name=th["font_name"], size=th.get("font_size_body", 11), bold=True,
                  color="000000")
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c


def _text_units(s):
    """열너비 단위 기준 문자열 폭: 한글·전각=2, 그 외=1.1(여유)."""
    return sum(2.0 if ord(ch) > 0x2E7F else 1.1 for ch in str(s))


def fit_rows(ws, min_row=1, max_row=None, min_col=1, max_col=None, line_height=15.0, max_height=409.0):
    """[폐기 — 사용자 확정 2026-09-09] 행높이는 건드리지 않는다(엑셀 기본 행높이 그대로).
    예전에는 wrap 셀의 줄 수를 추정해 행높이를 명시했으나, 이제 아무 것도 하지 않는다(호환용 no-op).
    긴 문장은 wrap_text 만 켜고 열너비·병합으로 처리한다."""
    return None


# ──────────────────────────────────────────────────────────────
# 조서 헤더 (회사명·조서번호·작성자·검토자 정보표 + 제목 밴드)
# ──────────────────────────────────────────────────────────────
HEADER_STYLES = ("grid", "line", "block")


def write_header(ws, meta, title=None, style="grid", theme=DEFAULT_THEME, last_col=7, first_col=2):
    """계정조서 상단 헤더. 반환: 본문 첫 표의 header_row (그 위 1행은 공백).

    meta  : {"회사명","조서번호","기준일","작성자","작성일자","검토자","검토일자","상태"} — 없는 키는 빈칸
    title : 제목 밴드 문구. **생략(권장)** 하면 시트명이 들어간다 — 밴드에는 시트명만 쓴다
            (사용자 확정 2026-09-10, `write_title` 과 같은 규약. 문자열을 줘도 `title_only` 로 설명을 잘라 낸다)
    style : "grid"  = 워드 조서와 동일 — 라벨음영 격자 정보표(2행×3쌍)가 제목 위, 제목 밴드 아래에 본문 (권장)
            "line"  = 제목 밴드 + 그 아래 한 줄 메타(회사 · 조서 · 작성 · 검토, 회색 9pt, 하단 얇은선) — 최소형
            "block" = 제목 밴드 + 좌(회사·조서·기준일)/우(작성·검토·상태) 2단 정보블록 — 메모형
    last_col: 제목 밴드·정보표가 차지할 마지막 열(본문 표 폭과 일치시킨다)."""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    title = plain_punct(title_only(title, ws))
    meta = {k: plain_punct(v) for k, v in (meta or {}).items()}
    fn = th["font_name"]
    lab_fill = _fill(th["header_bg"])
    thin = Side(style="thin", color=th.get("border_in", "BFBFBF"))
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    g = lambda k: meta.get(k, "") or ""

    def title_band(row):
        bar = _fill(th.get("title_bg", th["header_bg"]))
        for col in range(first_col, last_col + 1):
            c = ws.cell(row=row, column=col); c.fill = bar
            c.font = Font(name=fn, size=th["font_size_title"], bold=True,
                          color=th.get("title_font_color", th["header_font_color"]))
        c = ws.cell(row=row, column=first_col, value=title)
        c.alignment = Alignment(horizontal="left", vertical="center")

    def cell(row, col, value, label=False, size=10, align="left"):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name=fn, size=size, bold=label, color=th.get("text_color"))
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = box
        if label:
            c.fill = lab_fill
        return c

    ws.column_dimensions[get_column_letter(1)].width = 2.0
    if style == "grid":
        pairs = [[("회사명", g("회사명")), ("조서번호", g("조서번호")), ("기준일", g("기준일"))],
                 [("작성자", g("작성자")), ("작성일자", g("작성일자")), ("검토자", g("검토자") + ("  /  " + g("검토일자") if g("검토일자") else ""))]]
        # 라벨·값 열: first_col 부터 6칸(라벨,값 ×3). 열이 부족하면 값 칸을 마지막 열까지 병합
        cols = list(range(first_col, first_col + 6))
        for i, row_pairs in enumerate(pairs):
            r = 2 + i
            for j, (lab, val) in enumerate(row_pairs):
                cell(r, cols[2 * j], lab, label=True)
                cell(r, cols[2 * j + 1], val)
            for col in range(cols[-1] + 1, last_col + 1):   # 남는 열은 격자만 연장
                cell(r, col, None)
            if last_col > cols[-1]:
                ws.merge_cells(start_row=r, start_column=cols[-1], end_row=r, end_column=last_col)
        title_band(5)
        return 7
    if style == "line":
        title_band(2)
        parts = [g("회사명"), "조서 " + g("조서번호") if g("조서번호") else "", "기준일 " + g("기준일") if g("기준일") else "",
                 "작성 " + g("작성자") + (" " + g("작성일자") if g("작성일자") else ""),
                 "검토 " + g("검토자") + (" " + g("검토일자") if g("검토일자") else "")]
        c = ws.cell(row=3, column=first_col, value="  ·  ".join(x for x in parts if x.strip()))
        c.font = Font(name=fn, size=9, color=th.get("subtle_color", "595959"))
        c.alignment = Alignment(horizontal="right", vertical="center")
        ws.merge_cells(start_row=3, start_column=first_col, end_row=3, end_column=last_col)
        for col in range(first_col, last_col + 1):
            ws.cell(row=3, column=col).border = Border(bottom=thin)
        return 5
    if style == "block":
        title_band(2)
        left = [("회사명", g("회사명")), ("조서번호", g("조서번호")), ("기준일", g("기준일"))]
        right = [("작성자 / 일자", (g("작성자") + "  " + g("작성일자")).strip()),
                 ("검토자 / 일자", (g("검토자") + "  " + g("검토일자")).strip()), ("상태", g("상태"))]
        rc = max(first_col + 3, last_col - 1)
        for i in range(3):
            cell(4 + i, first_col, left[i][0], label=True); cell(4 + i, first_col + 1, left[i][1])
            cell(4 + i, rc, right[i][0], label=True); cell(4 + i, rc + 1, right[i][1])
        return 8
    raise ValueError("style must be one of " + str(HEADER_STYLES))


COVER_FIELDS = ("회사명", "결산일", "조서번호", "작성자")
COVER_SHEET = "표지"


def write_cover_sheet(wb, meta, title, theme=DEFAULT_THEME, index=0):
    """[legacy — 2026-09-09 폐기] 감사조서는 표지 시트를 두지 않고 시트마다 write_audit_header 를 쓴다. 기존 파일 호환용으로만 남긴다.
    (구) 조서 헤더 전용 '표지' 시트(2026-09-07): 워크북 맨 앞에 회사명·결산일·조서번호·작성자
    4항목만 담는다. 본문 시트는 열너비 영향을 받지 않도록 제목 밴드(B2)만 두고 헤더 정보를 넣지 않는다.
    meta: {"회사명","결산일","조서번호","작성자"} — 그 외 키는 무시. 반환: 표지 워크시트."""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    fn = th["font_name"]
    ws = wb.create_sheet(COVER_SHEET, index)
    ws.column_dimensions["A"].width = 2.0
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 60
    bar = _fill(th.get("title_bg", th["header_bg"]))
    for col in (2, 3):
        c = ws.cell(row=2, column=col); c.fill = bar
        c.font = Font(name=fn, size=th["font_size_title"], bold=True,
                      color=th.get("title_font_color", th["header_font_color"]))
    ws["B2"] = title
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color=th.get("border_in", "BFBFBF"))
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, k in enumerate(COVER_FIELDS):
        r = 4 + i
        a = ws.cell(row=r, column=2, value=k)
        a.font = Font(name=fn, size=th.get("font_size_body", 11), bold=True, color=th.get("text_color"))
        a.fill = _fill(th["header_bg"]); a.border = box
        a.alignment = Alignment(horizontal="left", vertical="center")
        b = ws.cell(row=r, column=3, value=(meta or {}).get(k, "") or "")
        b.font = Font(name=fn, size=th.get("font_size_body", 11), color=th.get("text_color"))
        b.border = box; b.alignment = Alignment(horizontal="left", vertical="center")
    set_workbook_font(ws, fn, th.get("font_size_body", 11))
    return ws


# 제목 밴드에서 잘라 내는 설명 구분자(사용자 확정 2026-09-10) — 시트명 뒤에 붙은 부제·설명을 뗀다
_TITLE_TAIL = re.compile(r"\s*(?:[—–\-|:·]|\(|\[|/)\s.*$|\s*[—–|]\s*.*$")


def sheet_label(ws):
    """시트 이름에서 앞 순번을 뗀 이름. "10 개요" → "개요".
    제목 밴드(write_title)·조서 헤더의 '시트명' 기본값이 모두 이 값이다."""
    return re.sub(r"^\d+\s+", "", ws.title).strip()


def title_only(text, ws=None):
    """제목 밴드에 넣을 문자열을 시트명 한 덩어리로 줄인다(사용자 확정 2026-09-10).
    - None·빈 값 → 시트명(sheet_label)
    - "20 연령분석 — 매출채권 연령별 잔액 검토" → "연령분석" (앞 순번·뒤 설명 제거)"""
    if text is None or not str(text).strip():
        return sheet_label(ws) if ws is not None else ""
    t = re.sub(r"^\d+\s+", "", str(text)).strip()
    t = _TITLE_TAIL.sub("", t).strip()
    return t or (sheet_label(ws) if ws is not None else str(text).strip())


def write_title(ws, title=None, theme=DEFAULT_THEME, last_col=7, first_col=2, tab=None):
    """본문 시트 제목 밴드만(B2, 헤더 정보 없음). 반환: 첫 표 header_row(4).

    **제목 밴드에는 시트명만 쓴다(사용자 확정 2026-09-10).** 회사명·기준일·목적 설명 같은 문구를
    제목에 덧붙이지 않는다. 설명이 필요하면 밴드 아래 `write_caption` 으로 따로 적는다.
    title : 생략(권장) 하면 시트명(앞 순번을 뗀 `ws.title`)이 그대로 들어간다.
            문자열을 주더라도 `title_only` 로 뒤의 설명(— – | : · ( [ / 뒤)을 잘라 낸다.
            (조서 헤더 6항목을 쓰는 `audit` 테마는 이 밴드를 쓰지 않는다 — 그쪽은 그대로.)
    tab: 시트 탭 색 역할(guide/output/calc/input/pbc/raw) — 주면 set_tab 을 같이 호출한다."""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    title = plain_punct(title_only(title, ws))
    ws.column_dimensions[get_column_letter(1)].width = 2.0
    bar = _fill(th.get("title_bg", th["header_bg"]))
    for col in range(first_col, last_col + 1):
        c = ws.cell(row=2, column=col); c.fill = bar
        c.font = Font(name=th["font_name"], size=th["font_size_title"], bold=True,
                      color=th.get("title_font_color", th["header_font_color"]))
    c = ws.cell(row=2, column=first_col, value=title)
    c.alignment = Alignment(horizontal="left", vertical="center")
    if tab is not None:
        set_tab(ws, tab, th)
    return 4


def set_tab(ws, role, theme=DEFAULT_THEME):
    """시트 탭 색을 역할에 맞춰 칠한다 (사용자 확정 2026-09-10).

    role : guide(안내·표지, 색 없음) / output(결과·총괄) / calc(계산·집계)
           / input(입력—작성자) / pbc(입력—회사 제공) / raw(원본·참고)
           한글·별칭도 받는다(TAB_ROLE_ALIASES): "회사"·"input_ext" → pbc 등.

    진하기가 자리를 말한다 — output 이 가장 진하고 raw 가 가장 연하다.
    탭 색은 시트의 **성격**에만 쓴다. 진행 상태(완료·검토중)나 임시 표시에 쓰지 않는다.
    한 파일에서 같은 역할은 같은 색이며 예외를 두지 않는다.
    반환: 칠한 hex (guide 는 None).
    """
    th = theme if isinstance(theme, dict) else get_theme(theme)
    key = TAB_ROLE_ALIASES.get(role, role)
    tabs = th.get("tab_colors", TAB_COLORS)
    if key not in tabs:
        raise KeyError(f"알 수 없는 탭 역할 '{role}'. 사용 가능: {', '.join(tabs)}")
    col = tabs[key]
    if col:
        ws.sheet_properties.tabColor = col
    else:
        ws.sheet_properties.tabColor = None
    return col


AUDIT_HEADER_FIELDS = ("회사명", "조서번호", "시트명", "결산일", "조서명", "작성자")


def write_audit_header(ws, meta, sheet_title=None, theme="audit", first_col=2, tab=None):
    """감사조서 시트 상단 조서 헤더(사용자 확정 2026-09-08). **감사조서 시트마다 반드시** 두고, 시트 제목 밴드는 두지 않는다.
    2행×3쌍(라벨|값), 병합 없음, B2:G3:
        [회사명|값][조서번호|값][시트명|값]  /  [결산일|값][조서명|값][작성자|값]
    - 조서번호: 시트번호 포함(예 "6000B-10") · 시트명: 시트 이름에서 앞 순번을 뺀 것(기본, "10 개요"→"개요") · 조서명: 조서(파일) 이름.
    - 서식: 색 없음(라벨 굵게만), 위(2행)·아래(3행) medium, 가운데 thin, 좌우 외곽선 없음, 11pt.
    - **열너비를 헤더 값에 맞춰 늘리지 않는다**(긴 조서명은 옆 칸으로 흘러도 둔다). 열너비는 본문 표 기준(autofit 상한 MAX_COL_WIDTH).
    반환: 첫 표 header_row(5) — 4행은 공백. meta: {"회사명","조서번호","결산일","조서명","작성자"}."""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    fnm = th["font_name"]; sz = th.get("font_size_body", 11)
    thin = Side(style="thin", color=th.get("border_color", "BFBFBF"))
    thick = Side(style="medium", color=th.get("border_color_outer", "000000"))
    last_col = first_col + 5
    def box_for(r, col):
        return Border(left=(None if col == first_col else thin), right=(None if col == last_col else thin),
                      top=(thick if r == 2 else thin), bottom=(thick if r == 3 else thin))
    m = {k: plain_punct(v) for k, v in (meta or {}).items()}
    # 시트명 기본값: ws.title 에서 앞의 순번("10 ", "20 ")을 뗀 이름(사용자 확정 2026-09-09) — "10 개요" → "개요"
    m.setdefault("시트명", sheet_title if sheet_title is not None else sheet_label(ws))
    g = lambda k: m.get(k, "") or ""
    ws.column_dimensions[get_column_letter(1)].width = 2.0
    rows = [AUDIT_HEADER_FIELDS[:3], AUDIT_HEADER_FIELDS[3:]]
    for i, labels in enumerate(rows):
        r = 2 + i
        for j, lab in enumerate(labels):
            c = ws.cell(row=r, column=first_col + 2 * j, value=lab)
            c.font = Font(name=fnm, size=sz, bold=True, color=th.get("text_color"))
            c.border = box_for(r, first_col + 2 * j)
            c.alignment = Alignment(horizontal="left", vertical="center")
            v = ws.cell(row=r, column=first_col + 2 * j + 1, value=g(lab))
            v.font = Font(name=fnm, size=sz, color=th.get("text_color"))
            v.border = box_for(r, first_col + 2 * j + 1)
            v.alignment = Alignment(horizontal="left", vertical="center")
    if tab is not None:
        set_tab(ws, tab, th)
    return 5


def write_section_bar(ws, row, text, theme=DEFAULT_THEME, first_col=2, last_col=8):
    """시트 소제목·섹션 바(표준조서의 "1. 총괄표"·"2. Test 수행" 줄): title_bg 색 밴드 + 흰 굵은 글씨를
    (audit 테마의 소제목 규약: 섹션 소제목을 **메인헤더 밴드(title_bg #393939 + 흰 글씨)** 로 — default 와 다른 점)
    first_col~last_col 에 깔고 텍스트를 첫 칸에 쓴다. 반환: 다음 표 header_row 권장값(row+2).
    (audit·default=진회색 #393939, procpa=딥네이비)"""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    bar = _fill(th.get("title_bg", th["header_bg"]))
    font = Font(name=th["font_name"], size=th.get("font_size_body", 11), bold=True,
                color=th.get("title_font_color", th["header_font_color"]))
    for col in range(first_col, last_col + 1):
        c = ws.cell(row=row, column=col); c.fill = bar; c.font = font
        c.alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=row, column=first_col, value=plain_punct(text))
    return row + 2


def write_tickmark(ws, cell, mark, theme=DEFAULT_THEME, note=None):
    """감사 틱마크(R=전기조서, GL=총계정원장, F=Footing, PR=제시재무제표, Ref.=조서번호 등)를
    표준조서 관행대로 **빨강(accent) 글씨·가운데 정렬**로 쓴다. note 를 주면 셀 메모로 설명을 붙인다."""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    c = ws[cell]
    c.value = mark
    c.font = Font(name=th["font_name"], size=th.get("font_size_body", 11),
                  color=th.get("accent_color", "FF0000"))
    c.alignment = Alignment(horizontal="center", vertical="center")
    if note:
        from openpyxl.comments import Comment
        c.comment = Comment(note, "audit")
    return c


def _fill_ranges(ws, hexcolor, ranges):
    bg = _fill(hexcolor)
    for rng in ranges:
        obj = ws[rng]
        if isinstance(obj, tuple):          # 범위 → 행 튜플들
            for rowcells in obj:
                for c in rowcells:
                    c.fill = bg
        else:                               # 단일 셀
            obj.fill = bg


def mark_cells(ws, theme, *, input=(), linked=(), todo=()):
    """셀 종류를 색으로 표시 (DCF·감사조서 관행).

      input  : 직접 입력(하드코딩) 셀 → note_fill(#FFFBEF, 크림)
      linked : 타시트 연결 셀        → linked_fill(#CCECFF, 파랑)
      todo   : 미입수 - 아직 회사에서 못 받은 자료 → todo_fill(#FFFF00, 노랑)
      (수식 셀은 별도 표시 없음 = 검정 본문)

        mark_cells(ws, "default", input=["C5:C9", "E12"], todo=["I5:I40"])
    """
    th = theme if isinstance(theme, dict) else get_theme(theme)
    if isinstance(input, str):
        input = [input]
    if isinstance(linked, str):
        linked = [linked]
    if input:
        _fill_ranges(ws, th.get("note_fill", "FFFBEF"), input)
    if isinstance(todo, str):
        todo = [todo]
    if linked:
        _fill_ranges(ws, th.get("linked_fill", "CCECFF"), linked)
    if todo:
        _fill_ranges(ws, th.get("todo_fill", "FFFF00"), todo)
    for rng in list(input) + list(linked) + list(todo):   # 빈 칸도 테마 글꼴(미입수 칸은 보통 비어 있다 - theme_lint '테마 외 폰트' 방지)
        obj = ws[rng]
        for c in ([x for r in obj for x in r] if isinstance(obj, tuple) else [obj]):
            if c.value is None:
                c.font = Font(name=th["font_name"], size=th["font_size_body"])


def highlight_hardcoded(ws, theme, *ranges):
    """하드코딩(입력) 셀을 #FFFBEF 로 표시. (mark_cells(input=...) 의 별칭)"""
    th = theme if isinstance(theme, dict) else get_theme(theme)
    _fill_ranges(ws, th.get("note_fill", "FFFBEF"), ranges)


# ──────────────────────────────────────────────────────────────
# 워크북 기본폰트 — 값 없는 빈 셀까지 테마 폰트로 통일
# ──────────────────────────────────────────────────────────────
def set_workbook_font(ws, name, size=11):
    """워크북 Normal 스타일 폰트를 바꿔, 값이 없는 셀도 지정 폰트로 렌더되게 한다.
    (openpyxl 기본 minorFont=Calibri 때문에 빈 셀이 Calibri로 보이는 문제 해결)"""
    try:
        wb = ws.parent
        for st in wb._named_styles:
            if getattr(st, "name", "") == "Normal":
                st.font = Font(name=name, size=size)
                return
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────
# 통합 진입 함수
# ──────────────────────────────────────────────────────────────
def apply_theme(ws, theme=DEFAULT_THEME, header_row=1, data_range=None,
                title_cell=None, number_format_cols=None,
                zebra=False, freeze=False, autofit=True, highlight_neg=None,
                margin=True):
    """워크시트에 테마를 한 번에 적용한다.

    ws                : openpyxl Worksheet
    theme             : 테마 이름(str) 또는 토큰 dict
    header_row        : 헤더가 있는 행 번호 (1-based)
    data_range        : 표 영역 "A1:E20". None 이면 사용 영역 자동 감지.
    title_cell        : 시트 제목 셀 "B2" (선택). 제목 색은 데이터 열까지 확장된다.
    number_format_cols: {"C": "thousands", ...} 열별 숫자 서식 (선택)
    zebra/autofit/highlight_neg : 각 규칙 on/off
    freeze            : 틀 고정 (기본 규칙: 끔). 필요할 때만 True.
    margin            : A열·1행을 여백으로 비움 (기본 규칙: A열 너비 2.0. 행높이는 지정하지 않는다 — 2026-09-09).
    """
    th = theme if isinstance(theme, dict) else get_theme(theme)
    # 워크북 기본폰트(Normal)를 테마 폰트로 → 값 없는 빈 셀까지 동일 폰트(맑은 고딕 등)
    set_workbook_font(ws, th["font_name"], th.get("font_size_body", 11))
    min_col, min_row, max_col, max_row = _bounds(ws, data_range)

    # 음수 강조(굵게)는 테마 설정(excel_highlight_neg)을 따른다 — 미지정 시.
    # (False면 숫자서식 [Red]만 적용 → 기본 빨강·보통굵기)
    if highlight_neg is None:
        highlight_neg = th.get("excel_highlight_neg", False)

    # 기본 규칙: A열·1행을 여백으로 비워둔다 (A열 너비 2.0). 행높이는 어디서도 지정하지 않는다(엑셀 기본값).
    if margin:
        ws.column_dimensions[get_column_letter(1)].width = 2.0

    # 제목 (표 영역 밖일 수 있으므로 먼저)
    if title_cell:
        col_letter, trow = coordinate_from_string(title_cell)
        tcol = column_index_from_string(col_letter)
        tc = ws[title_cell]
        if th.get("title_fill"):
            # 제목 바 색(title_bg)을 데이터가 있는 마지막 열까지 가로로 확장
            bar = _fill(th.get("title_bg", th["header_bg"]))
            for col in range(tcol, max_col + 1):
                bc = ws.cell(row=trow, column=col); bc.fill = bar
                bc.font = Font(name=th["font_name"], size=th["font_size_title"], bold=True,
                               color=th.get("title_font_color", th["header_font_color"]))
            tc.font = Font(name=th["font_name"], size=th["font_size_title"],
                           bold=True,
                           color=th.get("title_font_color", th["header_font_color"]))
        else:
            tc.font = Font(name=th["font_name"], size=th["font_size_title"],
                           bold=True, color=th["title_color"])
        tc.alignment = Alignment(horizontal="left", vertical="center")
        tc.value = plain_punct(tc.value)

    # 문장부호 규칙(2026-09-11): 표 영역 텍스트의 — → "-" (수식·숫자 제외)
    normalize_punct(ws, min_row, min_col, max_row, max_col)

    # 본문 → 줄무늬 → 헤더 → 테두리 순서 (뒤가 앞을 덮어씀)
    style_body(ws, th, header_row, min_col, min_row, max_col, max_row)
    if zebra:
        zebra_stripes(ws, th, header_row, min_col, min_row, max_col, max_row)
    style_header_row(ws, th, header_row, min_col, max_col)
    apply_borders(ws, th, min_col, min_row, max_col, max_row)

    if number_format_cols:
        set_number_formats(ws, th, number_format_cols, min_row, max_row, header_row)
    if highlight_neg:
        body_range = "%s%d:%s%d" % (
            get_column_letter(min_col), header_row + 1,
            get_column_letter(max_col), max_row)
        highlight_negatives(ws, th, body_range)
    if freeze:
        freeze_below_header(ws, header_row, min_col)
    if autofit:
        autofit_columns(ws, min_col, min_row, max_col, max_row)

    return ws


# ──────────────────────────────────────────────────────────────
# 고수준 표 생성 (데이터 + 옵션 → 완성된 themed 표)
# ──────────────────────────────────────────────────────────────
def write_table(ws, headers, rows, theme=DEFAULT_THEME, *,
                title=None, number_cols=None, currency_cols=(), percent_cols=(),
                total_rows=(), subtotal_rows=(), section_rows=(),
                input_cells=(), linked_cells=(), todo_cells=(),
                title_cell="B2", header_row=4, start_col=2, secondary=False):
    """데이터를 받아 **한 번에** themed 표를 작성한다. 반환: 사용한 data_range.

    레이아웃: 제목=title_cell(기본 B2), 헤더=header_row(기본 4), 데이터=그 아래(B5~). A열 여백.
      headers       : 헤더 셀 리스트
      rows          : 2차원 리스트. None 셀은 비움(섹션 라벨 행 등).
      title         : B2 시트 제목 (None이면 제목 미작성)
      number_cols   : {열키: 서식키/문자열}. 열키=0-based 데이터열 인덱스 또는 'C' 같은 문자.
      currency_cols : 통화 열 인덱스들 → 테마 기본 통화서식(전 테마 accounting = 원 단위 회계서식).
      percent_cols  : 퍼센트 열 인덱스들 → percent_acct.
      total/subtotal/section_rows : 데이터 1-based 인덱스(rows 기준).
      input_cells / linked_cells / todo_cells : 입력=크림·연결=파랑·미입수=노랑 으로 표시할 범위/셀 (mark_cells).
      secondary     : True 면 서브표(2차 표) — 헤더·소계·합계 배경을 header2_bg 로 통일.

    예) write_table(ws, ["계정","당기","전기"], data, theme="default",
                    title="재무상태표", currency_cols=[1,2], total_rows=[len(data)],
                    input_cells=["C5:D6"])
    """
    th = theme if isinstance(theme, dict) else get_theme(theme)
    n_cols = len(headers)

    if title is not None:
        ws[title_cell] = title
    for j, h in enumerate(headers):
        ws.cell(row=header_row, column=start_col + j, value=h)
    data_start = header_row + 1
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if v is not None:
                ws.cell(row=data_start + i, column=start_col + j, value=v)
    last_row = data_start + len(rows) - 1
    data_range = "%s%d:%s%d" % (get_column_letter(start_col), header_row,
                                get_column_letter(start_col + n_cols - 1), last_row)

    # 숫자서식: currency/percent → number_cols 순으로 합치기 (열키는 문자로 정규화)
    def _col(key):
        return key if isinstance(key, str) else get_column_letter(start_col + key)
    fmts = {}
    cur_fmt = th.get("currency_format", "accounting")
    for k in currency_cols:
        fmts[_col(k)] = cur_fmt
    for k in percent_cols:
        fmts[_col(k)] = "percent_acct"
    for k, v in (number_cols or {}).items():
        fmts[_col(k)] = v

    apply_theme(ws, theme=th, header_row=header_row, data_range=data_range,
                title_cell=(title_cell if title is not None else None),
                number_format_cols=(fmts or None))

    def _srow(idx):
        return data_start + idx - 1
    mx = start_col + n_cols - 1
    if secondary:
        style_header_row(ws, th, header_row, start_col, mx, secondary=True)
    for idx in section_rows:
        style_subheader_row(ws, th, _srow(idx), start_col, mx)
    for idx in subtotal_rows:
        style_total_row(ws, th, _srow(idx), start_col, mx, fill=False, secondary=secondary)
    for idx in total_rows:
        style_total_row(ws, th, _srow(idx), start_col, mx, fill=True, secondary=secondary)

    if input_cells or linked_cells or todo_cells:
        mark_cells(ws, th, input=list(input_cells), linked=list(linked_cells), todo=list(todo_cells))

    return data_range
