# -*- coding: utf-8 -*-
"""kicpa_core - 한공회 표준조서 감사 스킬 팩(kicpa-jet · kicpa-analytical · kicpa-sap) 공통 코드.

journal  분개장 읽기(더존 표준 머리글 · ERP 내보내기형, 여러 파일 합치기)
tb       TB 읽기(일반 표 · xlaudit tb 폴더 · douzone_tb json) - 금액은 차변 +
coa      계정 분류(더존 3자리 계정 체계 기준)와 IFRS · K-GAAP 과목 골격 매핑
wp       조서 머리 · 기초정보 · 이름 정의 · 하이퍼링크 등 조서 공통 부품
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILLS = os.path.dirname(os.path.dirname(HERE))
THEME_DIR = os.path.join(SKILLS, "excel-theme")
if THEME_DIR not in sys.path:
    sys.path.insert(0, THEME_DIR)


class KicpaError(Exception):
    pass
