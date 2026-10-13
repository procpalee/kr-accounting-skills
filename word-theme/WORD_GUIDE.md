# 워드 테마 규격서 (WORD_GUIDE)

`word-theme` 스킬이 적용하는 색·폰트·문단/표 서식의 **사람이 읽는 규격**이다.
색/폰트의 원본은 **`word_themes.py`(WORD_THEMES)** 이며, 바꾸려면 그 파일을 수정한다.
(워드 전용 테마에 없는 이름은 `..\excel-theme\excel_theme.py` 팔레트로 폴백)

## 1. 문서 기본 (apply_base_style / new_doc)
- 본문(Normal): 테마 `font_name` (default=맑은 고딕 / procpa=Pretendard), `body` pt, `text` 색.
- 한글(동아시아) 글꼴을 별도 지정(`w:eastAsia`)하여 한글에도 폰트가 적용된다.
- 바닥글 가운데 페이지 번호(PAGE 필드), `secondary` 색 9pt.
- `new_doc(theme, doc_type)` 이 테마를 문서에 기억시켜 이후 헬퍼가 자동 상속한다.

## 2. 제목부 (add_title — 5종 스타일)
| style | 구성 | 용도 (doc_type) |
|---|---|---|
| `bar` | 제목(`title`pt, title_color) + 부제(`subtitle`pt, secondary) + primary 강조 바 2.5pt | 기본 — 일반 보고서 (`report`) |
| `side` | 좌측 세로 primary 바(3.5pt) + 제목(title-4) + 부제 | 간단한 메모·검토 (`memo`) |
| `band` | 전폭 primary 배경 밴드 + 자간 캡션(secondary_lt) + 흰 제목(title-2) + 부제(secondary_lt) | 제안서·트렌디 (`proposal`) |
| `center` | 페이지 88% 폭 가운데 + 상하 secondary_lt 가는 선 + 중앙 제목(title-4) + 메타 라인 | 공식 의견서·공문 (`opinion`) |
| `meta` | 제목(title-5, 좌) + 문서정보 표(small, 우: 라벨=헤더색) + 하단 primary 2pt 선(고정 줄높이로 간격 최소) | 조서 (`workpaper`) |

- 모든 스타일의 제목은 두 줄로 꺾일 때 **고정 줄높이(폰트×1.15)** 로 줄간격을 압축한다.
- `meta=` dict: `caption` 키는 캡션 라벨(band), 나머지는 메타 라인(center)/문서정보 표(meta)에 표기.
- 헤딩: Heading 1=`primary`, 2=`primary_dark`, 3=`secondary` (default/procpa는 셋 동일색).

## 3. 표 (add_table)
- 헤더행: `table_header_fill` 배경 + `table_header_text` 굵은 글씨 + 가운데정렬 (라이트 헤더).
- 삼선표: `table_sides=False`(좌우 외곽선 없음) + 상/하 `table_edge_sz`(2pt) + 헤더 하단 옅은 구분선.
- 정렬: 숫자/`number_cols` 지정 열은 우측, 텍스트는 좌측.
- **숫자 표기(엑셀 NUMBER_FORMATS 규약 재현)**: 음수 = `negative` 색 **괄호** `(1,200)`, 0 = 대시 `-`.
  포맷 키: `accounting`(기본) / `percent`(x.x%) / `percent_acct`(x.xx%) / `decimal` / `multiple`(x.xx).
- 합계행(`total_last_row=True`): `band` 배경 + 굵게 + 상단 구분선.
- 줄무늬(`zebra`)는 default/procpa 모두 꺼져 있다.

## 4. 강조 박스 (add_callout)
- `note` 배경 + 좌측 `accent` 굵은 선(3pt)의 1칸 표. 핵심 결론/주의 문구에 사용.
  (default=회색 `F2F2F2`+`595959` / procpa=블루 틴트 `EFF5FF`+`2563EB`)

## 5. 커스텀
- 색: `word_themes.py` 의 WORD_THEMES hex 수정 → `..\excel-theme\build_all.py` 로 미리보기 재생성.
- 문서유형 프리셋: `word_themes.py` 의 `DOC_TYPES`(title_style·caption) 수정.
- 제목부 샘플: `python examples/title_samples.py`.
