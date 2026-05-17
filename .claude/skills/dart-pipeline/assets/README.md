# assets/ — 결정적 처리 영역

이 폴더의 스크립트와 프로젝트 루트의 파이썬 모듈이 **계산을 수행하는 단일 진입점**이다. LLM은 이 코드를 **호출만 한다** — 내부 로직을 추측해서 다른 방식으로 재구현하지 않는다.

## 위치

| 파일 | 위치 | 역할 |
|---|---|---|
| `dart_client.py` | 프로젝트 루트 | DART API · 공시 목록 · PDF 다운로드 |
| `pdf_to_excel.py` | 프로젝트 루트 | PDF → 4시트 Excel (재무상태표/포괄손익/자본변동/현금흐름) |
| `consolidate.py` | 프로젝트 루트 | PDF별 Excel 묶음 → 시계열 통합 |
| `extractor.py` | 프로젝트 루트 | 위 셋을 묶은 CLI (`--harness` 모드 지원) |
| `validate.py` | **이 폴더** | 회계등식·합계 대사·누락 계정 검증 |

루트 모듈을 이 폴더로 옮기지 않은 이유: 사용자가 `py extractor.py "..."` 한 줄로 호출하는 기존 사용성을 유지하기 위함. import 경로도 단순.

## 호출 규약

```powershell
# 1) 수집 + 변환 + 통합 (harness 경로 사용)
py extractor.py --harness "회사명"

# 2) 산출물 검증
py .claude/skills/dart-pipeline/assets/validate.py output/YYYY-MM/<회사>_시계열_*.xlsx
```

## 수정이 필요할 때

- **계정 매핑 변경** (예: "용역수익"을 "매출액"으로 인정): 코드 직접 수정 X. `references/assumptions/<회사>/key_items.md` 에 오버라이드 작성.
- **새 양식 발견** (예: 페이지 첫 줄이 다른 회사): 우선 회사별 패치로 시도. 보편적 양식이면 그때 `pdf_to_excel.py` 본체 수정.
- **검증 항목 추가**: `validate.py` 에 함수 추가.
