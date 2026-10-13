# -*- coding: utf-8 -*-
"""
■ 테마 확인·수정용 통합 빌드 스크립트 ■

색을 바꾸고 싶으면:
  1) excel_theme.py 상단 "편집 영역"의 PALETTES hex 값을 고친다 (3개 앱의 단일 출처).
  2) 이 스크립트를 실행한다:  python build_all.py   (excel-theme 폴더에서)
  3) 생성된 미리보기 파일들을 열어 확인한다 (경로는 실행 끝에 출력).

하는 일:
  • palette_swatch.html  — 모든 팔레트의 색을 라벨과 함께 한 눈에 보여주는 시각 확인용
  • 엑셀/워드/PPT 각 스킬의 build_preview.py 실행 → 미리보기 문서 일괄 생성
  • 생성된 파일 목록을 출력
"""
import os
import sys
import subprocess

# 한글 출력이 콘솔 인코딩에 관계없이 깨지지 않도록
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))   # excel-theme
SKILLS = os.path.dirname(HERE)                       # ~/.claude/skills

sys.path.insert(0, HERE)
from excel_theme import PALETTES, DEFAULT_THEME as DEFAULT_PALETTE  # noqa: E402

# 스와치에 표시할 의미색 순서
SWATCH_ROLES = ["primary", "primary_dark", "secondary", "secondary_lt",
                "band", "subheader", "surface", "note",
                "accent", "negative", "text", "header_font"]


def build_swatch(out_path):
    """모든 팔레트 색을 칩으로 보여주는 HTML 생성."""
    parts = ["<!doctype html><meta charset='utf-8'>",
             "<title>Office 팔레트 스와치</title>",
             "<style>",
             "body{font-family:Pretendard,'맑은 고딕',sans-serif;margin:32px;background:#fafafa;color:#222}",
             "h1{font-size:22px} h2{margin-top:36px;font-size:17px}",
             ".row{display:flex;flex-wrap:wrap;gap:10px;margin:12px 0}",
             ".chip{width:150px;border:1px solid #ddd;border-radius:8px;overflow:hidden;background:#fff}",
             ".sw{height:64px} .meta{padding:7px 9px;font-size:12px;line-height:1.5}",
             ".role{font-weight:600} .hex{color:#666;font-family:monospace}",
             ".tag{display:inline-block;background:#eee;border-radius:4px;padding:1px 6px;font-size:11px;margin-left:6px}",
             "</style>",
             "<h1>Office 공유 팔레트 스와치</h1>",
             "<p>색을 바꾸려면 <code>excel_theme.py</code> 의 PALETTES hex 값을 수정하고 "
             "<code>python build_all.py</code> 를 다시 실행하세요. "
             f"현재 기본 테마: <b>{DEFAULT_PALETTE}</b></p>"]
    for name, pal in PALETTES.items():
        tag = " <span class='tag'>DEFAULT</span>" if name == DEFAULT_PALETTE else ""
        parts.append(f"<h2>{name} — {pal.get('label','')}{tag}</h2>")
        parts.append(f"<p class='hex'>font: {pal['font_name']} → {pal['font_fallback']}</p>")
        parts.append("<div class='row'>")
        for role in SWATCH_ROLES:
            hexv = pal["colors"].get(role)
            if not hexv:
                continue
            # 글자색은 배경 밝기로 결정
            r, g, b = int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16)
            fg = "#fff" if (0.299 * r + 0.587 * g + 0.114 * b) < 140 else "#222"
            parts.append(
                f"<div class='chip'><div class='sw' style='background:#{hexv};color:{fg};"
                f"display:flex;align-items:center;justify-content:center;font-size:12px'>#{hexv}</div>"
                f"<div class='meta'><span class='role'>{role}</span><br>"
                f"<span class='hex'>#{hexv}</span></div></div>")
        parts.append("</div>")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    print("  생성:", out_path)


def run_builder(skill, rel="examples/build_preview.py"):
    path = os.path.join(SKILLS, skill, rel)
    if not os.path.exists(path):
        print(f"  (건너뜀: {path} 없음)")
        return
    print(f"\n▶ {skill}")
    # 서브프로세스로 격리 실행 (스킬별 import 격리 + 콘솔 인코딩 회피). 자식에 UTF-8 강제.
    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, path], env=env)
    if r.returncode != 0:
        raise SystemExit(f"{skill} 빌더 실패 (returncode={r.returncode})")


def main():
    print("=" * 60)
    print("Office 테마 전체 빌드")
    print("=" * 60)
    print("\n[1/2] 팔레트 스와치 생성")
    build_swatch(os.path.join(HERE, "palette_swatch.html"))

    print("\n[2/2] 앱별 미리보기 생성")
    for skill in ("excel-theme", "word-theme", "ppt-theme"):
        run_builder(skill)

    print("\n" + "=" * 60)
    print("완료. 아래에서 결과를 확인하세요:")
    print("  색 확인 :", os.path.join(HERE, "palette_swatch.html"))
    print("  엑셀    :", os.path.join(SKILLS, "excel-theme", "examples"))
    print("  워드    :", os.path.join(SKILLS, "word-theme", "examples"))
    print("  PPT     :", os.path.join(SKILLS, "ppt-theme", "examples"))
    print("=" * 60)


if __name__ == "__main__":
    main()
