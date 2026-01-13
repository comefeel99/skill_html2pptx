---
description: PPTX 변환 개선 작업 진행 워크플로우 - 테스트 케이스 추가를 포함한 체계적인 개선 프로세스
---

# PPTX 변환 개선 워크플로우

이 워크플로우는 HTML → PPTX 변환 개선 작업 시 테스트 케이스 추가를 포함한 체계적인 프로세스를 안내합니다.

## 1. 이슈 분석
- 문제가 발생하는 테스트 디렉토리 확인 (test1, test2, test3, test4)
- 문제가 발생하는 슬라이드 번호 확인
- 문제 유형 파악 (텍스트 누락, 중복, 위치 오류, 스타일 오류 등)

## 2. 테스트 케이스 추가 (수정 전 필수!)
문제 수정 전에 반드시 테스트 케이스를 먼저 추가합니다:

// turbo
```bash
# 예시 - 실제 이슈에 맞게 수정하세요
node add_test_case.js --id [이슈ID] --desc "[문제 설명]" \
  --test [테스트디렉토리] --slide [슬라이드번호] --type [assertion타입] [추가옵션]
```

### Assertion 타입별 옵션:
- `text_count`: `--content "텍스트" --count 1` (중복 방지)
- `text_exists`: `--content "텍스트"` (누락 방지)
- `text_not_exists`: `--content "텍스트"` (제거 확인)
- `similarity_threshold`: `--min 70` (품질 기준)
- `element_count`: `--count 5 --tolerance 1`
- `image_count`: `--count 3`

## 3. 회귀 테스트 실행 (수정 전 상태 확인)
// turbo
```bash
node run_regression.js
```
→ 새로 추가한 테스트 케이스가 FAIL인지 확인 (아직 수정 안했으므로)

## 4. 원인 분석 및 해결방안 제안
- 관련 코드 분석 (pptx/scripts/ 하위 파일들)
- 원인 분석 및 해결방안 정리
- **사용자 승인/선택 요청**: 코드를 수정하기 전에 분석된 원인과 해결방안을 사용자에게 공유하고, 개발 진행 허락 또는 옵션 선택을 받습니다.

## 5. 코드 수정
- 사용자의 허락 혹은 선택된 방안에 따라 코드 수정 작업 진행

## 6. 단일 테스트 실행
// turbo
```bash
node run_test.js test_data/[테스트디렉토리]
```

## 7. 회귀 테스트 재실행 (수정 후 검증)
// turbo
```bash
node run_regression.js
```
→ 새로 추가한 테스트 케이스가 PASS로 변경되었는지 확인
→ 기존 테스트 케이스들이 여전히 PASS인지 확인 (사이드 이펙트 검사)

## 8. 전체 테스트 실행 (최종 검증)
```bash
./execute_all_tests.sh
```

## 9. 완료
- 모든 테스트 PASS 확인
- 필요시 커밋

---

## 빠른 참조: 테스트 케이스 추가 예시

```bash
# 텍스트 중복 방지
node add_test_case.js --id issue_price_dup --desc "가격 중복 방지" \
  --test test3 --slide 7 --type text_count --content "5,400엔" --count 1

# 텍스트 누락 방지
node add_test_case.js --id issue_year_missing --desc "연도 누락 방지" \
  --test test3 --slide 1 --type text_exists --content "2024"

# 불필요 텍스트 제거 확인
node add_test_case.js --id issue_ghost_text --desc "고스트 텍스트 제거" \
  --test test2 --slide 5 --type text_not_exists --content "undefined"

# 슬라이드 품질 기준
node add_test_case.js --id quality_slide5 --desc "슬라이드5 품질" \
  --test test1 --slide 5 --type similarity_threshold --min 75
```
