---
name: dart-pipeline
description: DART 감사보고서 PDF에서 재무제표를 추출하여 시계열 Excel로 만드는 파이프라인. 사용 시점 — "감사보고서 추출", "회사명 재무 시계열", "DART 자동 다운로드", "재무제표 통합", "회계등식 검증" 류 요청. 비상장 외부감사 회사·상장사 비상장 시기 자료 분석에 사용.
---

# DART 감사보고서 파이프라인

## 역할 분담

| 영역 | 위치 | 성격 |
|---|---|---|
| 결정적 처리 | 프로젝트 루트의 `dart_client.py` / `pdf_to_excel.py` / `consolidate.py` / `extractor.py`, `assets/validate.py` | 파이썬 스크립트. LLM은 호출만, 내부 로직 수정·재구현 금지 |
| 가정·매핑 | `references/assumptions/<회사>/` | 회사별 KEY_ITEMS 오버라이드, 후보 선택 결정 등 |
| 산출물 스키마 | `references/schema/output_schema.md` | 시계열 Excel 시트·컬럼 표준 정의 |

## 표준 호출 순서

1. **수집**: `py extractor.py --harness "회사명"` — DART에서 PDF 자동 다운로드 → `_input/raw/<회사>/`, 변환 → `_workspace/<회사>/excels/`, 시계열 통합 → `output/YYYY-MM/<회사>_시계열_YYYY-MM-DD.xlsx`
2. **검증**: `py .claude/skills/dart-pipeline/assets/validate.py output/YYYY-MM/<회사>_시계열_*.xlsx` — 회계등식 + 누락 계정 체크. 결과는 `_workspace/<회사>/logs/validation_YYYY-MM-DD.log`
3. **분석**: 검증 통과한 시계열 데이터에 대해서만 텍스트 분석 작성

## 회사별 오버라이드 — 언제 만들고 어디에 두나

- 새 회사에서 회계등식이나 매출액 추출이 실패하면 `references/assumptions/<회사>/key_items.md` 작성.
- 형식은 `references/assumptions/_template.md` 참조.
- `consolidate.py`가 회사명 기준으로 이 파일을 자동 로드해서 KEY_ITEMS를 덮어쓴다.

## 금지

- `_input/raw/` 의 PDF를 어떤 방식으로든 수정·삭제·이동하지 않는다 (hook으로 강제됨).
- 회계 숫자를 LLM이 직접 계산하지 않는다. `validate.py`나 pandas 호출만 사용.
- `output/` 의 기존 산출물을 덮어쓰지 않는다. 새 타임스탬프로 별도 저장.
