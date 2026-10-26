# -*- coding: utf-8 -*-
"""분 · 반기 DSD 작성 한 번에: 밀기(roll-interim) → 채우기(fill) → 대사(recon) → 검산(check).

  python interim.py --base 반기최종.dsd --prior-same 전년3분기.dsd --workpaper 8110.xlsx \
                    [--prior-annual 전기감사보고서.dsd] [--config dsd_recon.json] [--out-dir 폴더] [--dart]

결과(out-dir, 기본 = 조서 폴더):
  <이름>_당전기밀기.dsd/.xlsx  밀기 결과 + '롤포워드' 시트(확인필요 · 미처리)
  <이름>_채움.dsd               당기 칸을 조서 값으로 채운 DSD  ← DART 편집기에서 열어 확인할 파일
  <이름>_채움_채움결과.xlsx     못 채운 칸 목록(조서에 없는 값 · 후보 여럿 · 전기 값 다름)
  <이름>_채움_조서대사.xlsx     recon 결과(확인 필요 시트)
  <이름>_채움_검증.xlsx         check 결과(합계 · 대사 · 표기)
dsdxlsx 는 PATH 의 `dsdxlsx` 명령(uv tool install) 또는 환경변수 DSDXLSX_HOME(저장소 폴더 - uv run)으로 찾는다.
"""
import argparse
import os
import shutil
import subprocess
import sys


def dsdxlsx_cmd():
    home = os.environ.get("DSDXLSX_HOME")
    if home and os.path.isdir(home):
        return ["uv", "run", "--no-editable", "--project", home, "dsdxlsx"], home
    exe = shutil.which("dsdxlsx")
    if exe:
        return [exe], None
    sys.exit("dsdxlsx 를 찾을 수 없습니다 - `uv tool install git+https://github.com/procpalee/dsdxlsx` "
             "또는 환경변수 DSDXLSX_HOME 에 저장소 폴더를 지정하세요")


def run(cmd, cwd, label, ok=(0,)):
    print(f"\n== {label}")
    r = subprocess.run(cmd, capture_output=True, cwd=cwd,
                       env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1"))
    out = (r.stdout or b"").decode("utf-8", "replace")
    lines = [ln for ln in out.splitlines() if ln.strip() and not ln.lstrip().startswith(("Building", "Built", "Installed", "Uninstalled"))]
    print("\n".join(lines[-25:]))
    if ok is not None and r.returncode not in ok:
        print((r.stderr or b"").decode("utf-8", "replace")[-1500:])
        sys.exit(f"{label} 실패(종료 코드 {r.returncode})")
    return out


def main():
    ap = argparse.ArgumentParser(description="분 · 반기 DSD 작성: 밀기 → 채우기 → 대사 → 검산")
    ap.add_argument("--base", required=True, help="뼈대 = 직전 중간보고서 최종 DSD (3분기 작업이면 반기 최종)")
    ap.add_argument("--prior-same", help="전년 동기 DSD (손익 · 현금흐름 · 흐름형 주석의 전기 값)")
    ap.add_argument("--prior-annual", help="전기 감사보고서 DSD (3분기 → 다음 해 1분기일 때 전기말 값)")
    ap.add_argument("--workpaper", required=True, help="당기 값 조서(8110 최종분석적검토 · PA 정산표 · 연결정산표) - 계산값 저장된 파일")
    ap.add_argument("--config", help="건별 설정 JSON(recon · fill 공용 - 시트 지정 · 빼는 칸 · 이름 다른 행)")
    ap.add_argument("--out-dir", help="결과 폴더 (기본: 조서 폴더)")
    ap.add_argument("--name", help="결과 파일 이름 앞부분 (기본: 뼈대 DSD 이름)")
    ap.add_argument("--dart", action="store_true", help="OpenDART 로 본 재무제표 전기 값 대사(상장사 - OPENDART_API_KEY)")
    a = ap.parse_args()
    cmd, cwd = dsdxlsx_cmd()
    out_dir = os.path.abspath(a.out_dir or os.path.dirname(os.path.abspath(a.workpaper)))
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.join(out_dir, a.name or os.path.splitext(os.path.basename(a.base))[0])
    roll = [*cmd, "roll-interim", os.path.abspath(a.base), "-o", stem + "_당전기밀기"]
    if a.prior_same:
        roll += ["--prior-same", os.path.abspath(a.prior_same)]
    if a.prior_annual:
        roll += ["--prior-annual", os.path.abspath(a.prior_annual)]
    if not a.dart:
        roll.append("--no-dart")
    run(roll, cwd, "1. 밀기(roll-interim)", ok=(0, 1))          # 1 = 확인필요 항목 있음(정상)
    cfg = ["--config", os.path.abspath(a.config)] if a.config else []
    filled = stem + "_채움.dsd"
    run([*cmd, "fill", stem + "_당전기밀기.dsd", os.path.abspath(a.workpaper), *cfg, "-o", filled], cwd, "2. 채우기(fill)")
    run([*cmd, "recon", filled, os.path.abspath(a.workpaper), *cfg, "-o", stem + "_채움_조서대사.xlsx"], cwd,
        "3. 대사(recon)")
    run([*cmd, "check", filled, "-o", stem + "_채움_검증.xlsx"], cwd, "4. 검산(check)", ok=None)   # 검출 수가 종료 코드
    print(f"\n다음: {filled} 를 DART 편집기에서 열어 '못 채움' · '확인 필요' · '롤포워드 확인필요'를 정리")


if __name__ == "__main__":
    main()
