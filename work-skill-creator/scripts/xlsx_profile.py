"""엑셀 파일 하나의 구조 점검 - 어떻게 고칠지(openpyxl / 엑셀 COM) 정하기 전에 돌린다.

보는 것
  시트별: 사용 범위 · 값 칸 · 수식 칸 · 배열수식 · 병합 · 데이터 유효성 · 조건부 서식 · 숨김 · 틀 고정
  통합문서: 이름 정의 · 외부 연결 · 표(ListObject) · 그림/차트 · 메모 · 스레드 메모 · 피벗 · 슬라이서 ·
           매크로(VBA) · 웹 추가기능 · 컨트롤
  → openpyxl 로 저장하면 사라지거나 깨질 수 있는 요소가 있으면 '엑셀(COM)로 고치기'를 권한다.
  수식 안에 박힌 기간 숫자(DATE(2026,…) · 202606 등)를 찾아 알려 준다 - 다음 기간에 사람이 손으로 고쳐 온 곳이다.
  --labels 를 주면 시트마다 머리글 · 항목 이름(이름표)을 보여 준다 - 코드가 '이름표로 위치 찾기'를 할 때 쓴다.

사용: python -X utf8 xlsx_profile.py <파일.xlsx> [--labels] [--max-labels 40]
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

# 수식 속 기간 숫자: DATE(2026,…) · 202606 · 20260630 · "2026-06-30" · "2026.06"
HARD_PERIOD = re.compile(r"DATE\(\s*20\d{2}|(?<![\w$])20\d{2}(0[1-9]|1[0-2])([0-3]\d)?(?![\d])|\"20\d{2}[-./](0[1-9]|1[0-2])")

# openpyxl 로 열고 저장할 때 잃거나 깨지는 것(경험상) - (zip 안 경로 패턴, 이름, 설명)
RISKY = [
    (r"xl/drawings/.*\.xml$", "그림 · 도형 · 차트", "도형 · 일부 차트 서식이 사라질 수 있음"),
    (r"xl/media/", "이미지", "openpyxl 이 다시 쓰면서 위치 · 크기가 바뀌거나 빠질 수 있음"),
    (r"xl/threadedComments/", "스레드 메모", "openpyxl 이 읽지 못해 사라짐"),
    (r"xl/comments\d*\.xml$", "메모", "모양(크기 · 위치)이 바뀜"),
    (r"xl/pivotTables/", "피벗 테이블", "새로 고침 정보가 깨질 수 있음"),
    (r"xl/slicers/|xl/slicerCaches/", "슬라이서", "사라짐"),
    (r"xl/timelines/", "시간 표시 막대", "사라짐"),
    (r"xl/vbaProject\.bin$", "매크로(VBA)", "keep_vba=True 로 .xlsm 저장하지 않으면 사라짐"),
    (r"xl/webextensions/", "웹 추가기능", "파일이 엑셀에서 안 열리는 사례 있음"),
    (r"xl/activeX/|xl/ctrlProps/", "컨트롤(체크박스 · 단추)", "사라지거나 동작이 끊김"),
    (r"xl/externalLinks/", "외부 연결", "연결 경로 · 캐시값이 바뀔 수 있음"),
    (r"customXml/", "사용자 XML(민감도 레이블 등)", "대개 유지되나 확인 필요"),
    (r"xl/model/", "데이터 모델(파워 피벗)", "사라짐"),
    (r"xl/connections\.xml$", "쿼리 연결(파워 쿼리)", "새로 고침 설정이 깨질 수 있음"),
]


def zip_scan(path: Path) -> dict:
    found = Counter()
    ext_targets = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for n in names:
            for rx, label, _ in RISKY:
                if re.search(rx, n):
                    found[label] += 1
        for n in names:
            if re.fullmatch(r"xl/externalLinks/_rels/externalLink\d+\.xml\.rels", n):
                xml = z.read(n).decode("utf8", "ignore")
                ext_targets += re.findall(r'Target="([^"]+)"', xml)
        ext_extra = []
        for n in names:  # 조건부 서식 · 유효성 확장(x14) - openpyxl 이 버린다
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n):
                xml = z.read(n).decode("utf8", "ignore")
                if "x14:conditionalFormatting" in xml:
                    ext_extra.append("조건부 서식 확장(x14)")
                if "x14:dataValidation" in xml:
                    ext_extra.append("데이터 유효성 확장(x14, 다른 시트 목록 참조)")
                if "x14:sparkline" in xml:
                    ext_extra.append("스파크라인")
        for e in set(ext_extra):
            found[e] += ext_extra.count(e)
    return {"found": found, "external": ext_targets}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="엑셀 파일 구조 점검")
    ap.add_argument("file")
    ap.add_argument("--labels", action="store_true", help="시트마다 앞쪽 글자 칸(이름표) 보기")
    ap.add_argument("--max-labels", type=int, default=40)
    a = ap.parse_args(argv)
    p = Path(a.file)
    if not p.exists():
        print(f"[오류] 파일이 없습니다: {p}", file=sys.stderr)
        return 2
    try:
        import openpyxl
    except ImportError:
        print("[오류] openpyxl 이 필요합니다", file=sys.stderr)
        return 2

    z = zip_scan(p)
    wb = openpyxl.load_workbook(p, data_only=False, keep_vba=p.suffix.lower() == ".xlsm")
    print(f"# 구조 점검 - {p.name}")
    print(f"시트 {len(wb.sheetnames)}개 · 이름 정의 {len(list(wb.defined_names))}개")
    print()
    print("| 시트 | 상태 | 사용 범위 | 값 칸 | 수식 칸 | 배열수식 | 병합 | 유효성 | 조건부서식 | 표 | 틀 고정 |")
    print("|---|---|---|---|---|---|---|---|---|---|---|")
    total_f = total_v = 0
    for ws in wb.worksheets:
        nv = nf = na = 0
        for (r, c), cell in getattr(ws, "_cells", {}).items():
            v = cell.value
            if v is None or v == "":
                continue
            if isinstance(v, str) and v.startswith("="):
                nf += 1
            elif getattr(v, "text", None) is not None:
                na += 1
                nf += 1
            else:
                nv += 1
        total_f += nf
        total_v += nv
        dv = len(ws.data_validations.dataValidation) if ws.data_validations else 0
        cf = len(list(ws.conditional_formatting))
        tables = len(getattr(ws, "tables", {}) or {})
        state = {"visible": "보임", "hidden": "숨김", "veryHidden": "아주 숨김"}.get(ws.sheet_state, ws.sheet_state)
        print(f"| {ws.title} | {state} | {ws.calculate_dimension()} | {nv:,} | {nf:,} | {na} | {len(ws.merged_cells.ranges)} "
              f"| {dv} | {cf} | {tables} | {ws.freeze_panes or ''} |")
    print()
    print(f"합계: 값 칸 {total_v:,} · 수식 칸 {total_f:,}"
          + (f" (수식 비율 {total_f / (total_f + total_v):.0%})" if total_f + total_v else ""))

    risky = [(label, n) for label, n in z["found"].items()]
    print()
    if risky:
        print("## openpyxl 로 저장하면 위험한 요소")
        for label, n in risky:
            desc = next((d for _, l, d in RISKY if l == label), "openpyxl 이 버림")
            print(f"- {label} ({n}) - {desc}")
        print()
        print("→ 이 파일은 **엑셀(COM)로 열어 고치는 것**을 권한다(wsc_xl.ExcelEdit). "
              "openpyxl 로 고쳐야 한다면 사본에서 시험하고, 엑셀로 열어 확인한 뒤 진행한다.")
    else:
        print("openpyxl 로 저장해도 잃는 요소가 보이지 않는다(그래도 저장 뒤 엑셀로 다시 계산 · 확인).")
    if z["external"]:
        print()
        print("## 외부 연결(다른 파일을 참조하는 수식)")
        for t in z["external"]:
            print(f"- {t}")
        print("→ 연결된 파일이 이번 기간 것으로 바뀌어야 하는지 확인한다(지난 기간 파일을 계속 보고 있으면 숫자가 옛 값).")

    names = list(wb.defined_names)
    if names:
        print()
        print("## 이름 정의(앞 20개)")
        for n in names[:20]:
            try:
                dn = wb.defined_names[n]
                print(f"- {n} = {dn.attr_text}")
            except Exception:
                print(f"- {n}")

    # 수식 안에 박힌 날짜 · 기간 숫자 - 다음 기간에 손으로 고쳐 온 곳일 수도, 계약 개시일처럼 고정 상수일 수도 있다
    hard_period, hard_date = [], []
    for ws in wb.worksheets:
        for (r, c), cell in getattr(ws, "_cells", {}).items():
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and HARD_PERIOD.search(v):
                item = f"{ws.title}!{cell.coordinate} {v[:60]}"
                (hard_date if re.search(r"DATE\(\s*20\d{2}", v) else hard_period).append(item)
    if hard_period or hard_date:
        print()
        print(f"## 수식 안에 박힌 기간 숫자 · 날짜 (기간 숫자 {len(hard_period)}칸 · DATE() 날짜 {len(hard_date)}칸)")
        for h in (hard_period[:8] + hard_date[:4]):
            print(f"- {h}")
        print("→ 이 중 기간마다 바뀌는 것만 기준일(기준 년월) 칸을 참조하게 바꾼다. 계약 개시일 같은 고정 날짜는 그대로 둔다.")
        print("  어느 것이 기간마다 바뀌었는지는 `compare_xlsx.py --period <그전> <지난>` 의 '수식 안에 박힌 기간 숫자가 바뀜' 으로 확인한다.")

    if a.labels:
        print()
        print("## 이름표(시트마다 머리글 · 항목 이름 - 명세 같은 데이터 행은 뺌)")
        for ws in wb.worksheets:
            rows = {}
            for (r, c), cell in getattr(ws, "_cells", {}).items():
                if cell.value not in (None, ""):
                    rows.setdefault(r, []).append(cell)
            labels = []
            for r in sorted(rows):
                cells = sorted(rows[r], key=lambda x: x.column)
                has_num = any(isinstance(x.value, (int, float, dt.date)) and not isinstance(x.value, bool) for x in cells)
                if len(cells) >= 3 and has_num:  # 거래처 · 금액이 든 데이터 행
                    continue
                for x in cells:
                    v = x.value
                    if isinstance(v, str) and v.strip() and not v.startswith("=") and not v.strip().isdigit():
                        labels.append(f"{x.coordinate} {v.strip()[:30]}")
                if len(labels) >= a.max_labels:
                    break
            print(f"- **{ws.title}**: " + " · ".join(labels[: a.max_labels]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
