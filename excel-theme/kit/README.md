# excel-theme 킷판 (강의·가이드 배포용)

클로드엑셀 강의·가이드(claude-excel-start) 실습 킷에 넣는 `excel-theme` 스킬의 **문서 원본**이다.
코드(서식 엔진)는 이 폴더에 두지 않고 상위 `excel-theme/` 원본을 그대로 가져와 조립한다.

- 테마: `default` 하나(정산표·결산 회색 표)
- 한 패키지로 두 경로: 클로드 앱·클로드 코드 = openpyxl(`excel_theme.py`), 클로드 엑셀 = Office.js(`make_code.py`)
- 색의 단일 출처 = `excel_theme.py` 편집 영역(PALETTES). `scripts/sync_theme.py` 가 `theme.json`·`references/themes.md` 를 다시 만든다
- 회사 서식으로 바꾸는 절차 = `references/customize.md` (실습 I 에서 skill-creator 가 따른다)

## 빌드

```bash
python -X utf8 tools/build_kit_skill.py        # → dist/excel-theme-kit/excel-theme.zip
```

빌드 단계: 원본 복사 → PALETTES·ALIASES 를 default 로 축소 → 동기화 → 검증 → ZIP.
검증 = openpyxl 결과와 Office.js ops 셀 단위 대조(`officejs/tests/check_parity.py`) · Node 목 실행(`check_officejs.py`) ·
`references/openpyxl.md` 예제 코드 실행 + `theme_lint` 위반 0 · 실습 I 모의(색만 바꾼 사본에서 두 엔진이 같이 바뀌는지, 동기화 누락을 `--check` 가 잡는지).
하나라도 실패하면 ZIP 을 만들지 않는다.

배포: 만든 ZIP 을 킷 zip(`procpa_portfolio/public/downloads/materials/claude-excel-guide-2026/claude-excel-guide-kit.zip`)의 `3부_스킬/excel-theme.zip` 으로 교체한다.
