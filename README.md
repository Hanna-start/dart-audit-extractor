# DART 감사보고서 → 재무제표 시계열 Excel (원-샷 자동화)

회사명 한 줄 입력하면 DART에서 모든 감사보고서/연결감사보고서 PDF를 자동으로 받아 → 재무제표 4종(BS/IS/SCE/CF) Excel로 변환 → 시계열 통합 Excel까지 만든다.

대상:
- **사업보고서 미제출 비상장 외부감사 회사** (DART OpenAPI `fnlttSinglAcntAll`은 정기보고서만 커버)
- 상장사의 비상장 시기 자료 (IPO 전 5~10년)

## 설치

```powershell
py -m pip install -r requirements.txt
```

DART API 키 발급: https://opendart.fss.or.kr (무료, 20,000 req/일)

```powershell
$env:DART_API_KEY = "40자리키"
```

## 사용

```powershell
# 기본: 전체 연도, 별도+연결 모두
py extractor.py "카카오스타일"

# 회사 검색만 (corp_code 후보 확인)
py extractor.py --search "쿠팡"

# 기간 한정 (공시 접수일 기준)
py extractor.py "카카오스타일" --from 2018 --to 2024

# 별도만 / 연결만
py extractor.py "카카오스타일" --only separate
py extractor.py "카카오스타일" --only consolidated

# PDF만 받고 변환 스킵
py extractor.py "카카오스타일" --download-only
```

회사명이 모호하면 (예: "쿠팡" → 쿠팡, 쿠팡주식회사, 쿠팡로지스틱스 등) 후보 목록이 뜨고 번호로 선택.

## 출력

```
data/
├─ CORPCODE.xml                          # 회사코드 캐시 (없으면 자동 다운)
└─ <회사명>/
    ├─ pdfs/                             # 다운받은 감사보고서 원본 PDF
    │   ├─ 감사보고서_2023.12_20240409000213.pdf
    │   ├─ 연결감사보고서_2023.12_20240409000216.pdf
    │   └─ ...
    ├─ excels/                           # PDF별 4시트 Excel
    │   ├─ 감사보고서_2023.12_20240409000213.xlsx
    │   └─ ...
    └─ <회사명>_시계열.xlsx               # 최종 통합 (시계열 분석용)
```

`<회사명>_시계열.xlsx` 시트 구성:

| 시트 | 내용 |
|---|---|
| `long_data` | 회사 × 연도 × 구분 × 재무제표 × 계정 × 값 (분석용 long format) |
| `별도_재무상태표` 등 8개 | 구분(별도/연결) × 재무제표 4종 피벗 (계정 × 연도) |
| `raw_*` | PDF별 원본 표 통합 (검증/탐색용) |

마지막 단계에서 회계등식(자산=부채+자본, 매출=원가+매출총이익)을 자동 검증하여 콘솔에 ✓/누락 표시.

## 동작 원리

1. **회사 검색**: DART `corpCode.xml`에서 키워드 매칭 → 후보 목록
2. **공시 조회**: `/api/list.json?pblntf_ty=F` 로 외부감사관련 공시 → 보고서명에 "감사보고서" 포함된 것만
3. **PDF 다운로드**: DART OpenAPI에 PDF 직접 다운 엔드포인트가 없어, 공시 뷰어 페이지(`dart.fss.or.kr/dsaf001/main.do`)에서 `openPdfDownload(rcept_no, dcm_no)` JS 호출의 dcm_no를 파싱한 뒤 `pdf/download/pdf.do` 호출. 세션 쿠키 + Referer + 브라우저 UA 필수.
4. **PDF → Excel**: `pdfplumber`로 페이지 첫 줄이 "재 무 상 태 표" 같은 큰 제목이면 새 섹션 시작 → 표 추출 → 컬럼 정규화(_num 컬럼 부착)
5. **시계열 통합**: 메타의 사업연도(보고기간 종료일 기준, NOT 공시일) + 핵심 계정 정규식 매칭 → long-format → 피벗

## 한계

- **스캔 PDF**: `pdfplumber`로 텍스트 추출 불가. OCR 별도 필요.
- **본 재무제표만**: 주석 미포함.
- **계정 매핑은 핵심 항목만**: 회사 간 비교를 정밀하게 하려면 `consolidate.py`의 `KEY_ITEMS` 정규식 확장 필요.
- **DART 페이지 구조 의존**: PDF 다운로드는 OpenAPI가 아닌 웹 페이지 경로 사용. DART가 페이지 구조를 바꾸면 `dart_client._DCM_PATTERNS` 등 조정 필요.

## 트러블슈팅

- **"dcm_no 추출 실패"**: 해당 공시가 PDF 첨부가 아닐 수 있음 (정정공시 일부, 첨부 없는 텍스트 공시). 보고서명 확인.
- **"PDF 응답 아님"**: User-Agent 차단 또는 rate limit. `dart_client.download_pdf`에서 UA를 확인하거나 잠시 후 재시도.
- **"법인세비용 None"**: K-GAAP "법인세등" 라벨이 매칭 안 됨. `consolidate.py`의 `KEY_ITEMS["포괄손익계산서"]` 정규식 확장.

## 라이브러리 사용

CLI 외에 모듈로도 사용 가능:

```python
from extractor import run_pipeline
out = run_pipeline("카카오스타일", api_key="...", year_from=2018, year_to=2024)
print(out)   # → data/카카오스타일/카카오스타일_시계열.xlsx
```

또는 단계별:

```python
import dart_client, pdf_to_excel, consolidate
dart_client.ensure_corpcode(key, "data/CORPCODE.xml")
cands = dart_client.find_companies("무신사", "data/CORPCODE.xml")
disclosures = dart_client.list_audit_disclosures(key, cands[0]["corp_code"])
for d in dart_client.dedupe_latest(disclosures):
    dart_client.download_pdf(d.rcept_no, f"data/pdfs/{d.rcept_no}.pdf")
# ... pdf_to_excel.extract_financials() / consolidate.run() ...
```
