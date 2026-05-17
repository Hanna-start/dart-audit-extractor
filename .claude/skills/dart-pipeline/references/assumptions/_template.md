# 회사별 KEY_ITEMS 오버라이드 템플릿

이 폴더(`assumptions/`)는 회사별 가정·매핑 결정을 시계열로 보존한다.

새 회사 추가 시 `references/assumptions/<회사명>/key_items.md` 생성 (회사명은 산출물 폴더명과 동일하게).

## 형식 (YAML front-matter + 설명)

```yaml
---
company: 카카오스타일
applied_from: 2026-05    # 이 가정이 처음 적용된 월
overrides:
  포괄손익계산서:
    - key: 매출액
      pattern: "매출액$|영업수익$|용역수익$|거래수수료수익$"
      transform: null
    - key: 매출원가
      pattern: "매출원가$|용역원가$"
      transform: abs
  재무상태표:
    - key: 자산총계
      pattern: "자산총계$"
      transform: null
---

# 카카오스타일 매핑 결정

## 왜 "거래수수료수익"을 매출액으로 인정했는가
플랫폼 회사라 매출 본질이 거래 중개 수수료. K-IFRS 손익계산서에 "거래수수료수익"으로 표시.
2026-05 추출 시 매출액 None 발생 → 본 매핑 추가.

## 검증
- 2024 별도: 매출 X,XXX억, 회계등식 ✓
- 2024 연결: 매출 X,XXX억, 회계등식 ✓
```

## 적용 방식

`consolidate.py`가 회사명을 보고 `references/assumptions/<회사>/key_items.md` 의 YAML front-matter를 읽어 기본 `KEY_ITEMS`를 덮어쓴다. 파일이 없으면 기본값 사용.

## 변경 이력

매핑을 바꾸면 새 파일을 같은 폴더에 추가하지 말고, 기존 파일을 수정한 뒤 git/manual로 변경 이력 관리. 큰 매핑 변경이면 `applied_from` 갱신.

## 금지

- 산출물 파일을 직접 손으로 수정해서 매핑 맞추지 마라. 반드시 이 파일을 갱신하고 재추출.
