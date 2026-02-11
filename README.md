# 방화벽 정책 검증 자동화 스크립트

이 프로젝트는 방화벽 정책 변경이 성공적으로 적용되었는지 자동으로 검증하는 스크립트입니다.

## 개요

방화벽 정책 파일은 DRM(Digital Rights Management)이 적용되어 있어 일반적인 Excel 라이브러리로는 읽을 수 없습니다. 이 스크립트는 `xlwings`를 사용하여 DRM 보호 파일을 처리하고, 정책 변경 사항을 검증합니다.

## 주요 기능

1. **정책 파일 파싱**: DRM이 적용된 Excel 파일에서 정책 정보 추출
   - `Rulename`: 정책 이름
   - `Enable`: 정책 활성화 상태

2. **성능 최적화**: 전체 범위를 한 번에 읽어 처리 속도 향상

3. **빈 셀 처리**: 정책 파일의 많은 빈 공간을 자동으로 건너뛰고 실제 데이터만 추출

4. **정책 검증**: Running 정책과 Candidate 정책을 비교하여 변경 사항 확인

## 워크플로우

```
┌─────────────────────────────────────────────────────────────┐
│ 1. 정책 파일 준비                                            │
│    - running_policy.xlsx (현재 실행 중인 정책)              │
│    - candidate_policy.xlsx (변경 예정 정책)                 │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 2. 정책 파일 파싱                                            │
│    - xlwings로 DRM 보호 파일 열기                           │
│    - 전체 데이터를 한 번에 읽기 (성능 최적화)                │
│    - 헤더 행 자동 탐지 ('Rulename', 'Enable')               │
│    - 빈 셀 제거 및 데이터 정제                               │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 3. 대상 정책 목록 로드                                       │
│    - 검증할 정책 이름 목록 읽기                              │
│    - 텍스트 파일(.txt) 또는 Excel 파일 지원                 │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 4. 정책 검증 (향후 구현)                                     │
│    - Running vs Candidate 비교                              │
│    - 대상 정책이 삭제되었는지 확인                           │
│    - 대상 정책이 비활성화되었는지 확인                       │
│    - 검증 결과 리포트 생성                                   │
└─────────────────────────────────────────────────────────────┘
```

## 사용 방법

### 기본 사용 (파싱만)

```bash
python parse_firewall_policy.py
```

이 명령은 `running_policy.xlsx`와 `candidate_policy.xlsx`를 파싱하여 
처리된 데이터를 Excel 파일로 저장합니다.

### 프로그래밍 방식 사용

```python
from parse_firewall_policy import parse_policy_file, load_target_policies, validate_policy_changes

# 정책 파일 파싱
running_df = parse_policy_file("running_policy.xlsx")
candidate_df = parse_policy_file("candidate_policy.xlsx")

# 대상 정책 목록 로드
target_policies = load_target_policies("target_policies.txt")

# 정책 변경 사항 검증
validation_results = validate_policy_changes(
    running_df, 
    candidate_df, 
    target_policies
)

# 결과 출력 및 저장
print(validation_results)
validation_results.to_excel("validation_report.xlsx", index=False)
```

### 파일 구조

```
validate-policy/
├── parse_firewall_policy.py      # 메인 스크립트
├── running_policy.xlsx            # 현재 실행 중인 정책 파일
├── candidate_policy.xlsx          # 변경 예정 정책 파일
├── target_policies.txt           # 검증할 정책 이름 목록 (선택사항)
└── README.md                     # 이 파일
```

## 함수 설명

### `parse_policy_file(file_path: str) -> pd.DataFrame`

정책 Excel 파일을 파싱하여 `Rulename`과 `Enable` 컬럼을 추출합니다.

**성능 최적화:**
- 전체 범위를 한 번에 읽어서 pandas DataFrame으로 변환
- 개별 셀 읽기 대신 벌크 읽기로 속도 향상
- 헤더 탐지는 DataFrame 연산으로 처리

**파라미터:**
- `file_path`: Excel 파일 경로 (DRM 보호 파일 가능)

**반환값:**
- `pd.DataFrame`: 'Rulename'과 'Enable' 컬럼을 가진 DataFrame

### `load_target_policies(file_path: str) -> list`

대상 정책 이름 목록을 파일에서 읽어옵니다.

**지원 형식:**
- 텍스트 파일 (.txt): 한 줄에 하나의 정책 이름
- Excel 파일 (.xlsx, .xls): 첫 번째 컬럼의 값들을 읽음 (벌크 읽기로 성능 최적화)

**파라미터:**
- `file_path`: 정책 이름 목록 파일 경로

**반환값:**
- `list`: 정책 이름 리스트

### `validate_policy_changes(running_df, candidate_df, target_policies) -> pd.DataFrame`

Running 정책과 Candidate 정책을 비교하여 대상 정책들의 변경 사항을 검증합니다.

**검증 항목:**
- `DELETED`: 정책이 삭제되었는지 확인 (Running에는 있지만 Candidate에는 없음)
- `DISABLED`: 정책이 비활성화되었는지 확인 (Enable 값이 Enabled → Disabled로 변경)
- `RE_ENABLED`: 정책이 다시 활성화되었는지 확인 (Disabled → Enabled)
- `NO_CHANGE`: 변경 없음
- `NOT_IN_RUNNING`: Running 정책에 존재하지 않음

**파라미터:**
- `running_df`: Running 정책 DataFrame (Rulename, Enable 컬럼)
- `candidate_df`: Candidate 정책 DataFrame (Rulename, Enable 컬럼)
- `target_policies`: 검증할 대상 정책 이름 리스트

**반환값:**
- `pd.DataFrame`: 검증 결과 리포트
  - 컬럼: `['Policy', 'Status', 'Running_Enable', 'Candidate_Enable', 'Message']`

## 요구사항

```txt
xlwings>=0.30.0
pandas>=1.5.0
openpyxl>=3.0.0  # Excel 파일 쓰기용
```

설치:
```bash
pip install xlwings pandas openpyxl
```

## 성능 개선 사항

### 이전 버전의 문제점
- 각 셀을 개별적으로 읽어 매우 느림 (`ws.range((row, col)).value`)
- 헤더 탐지를 위한 이중 루프로 인한 성능 저하
- 데이터 행도 하나씩 읽어 처리 시간 증가

### 개선된 버전
- ✅ 전체 범위를 한 번에 읽기 (`data_range.options(pd.DataFrame)`)
- ✅ pandas DataFrame 연산으로 헤더 탐지
- ✅ 벌크 처리로 처리 속도 대폭 향상

## 향후 개발 계획

- [ ] 정책 비교 기능 구현
- [ ] 삭제/비활성화 검증 로직 추가
- [ ] 검증 결과 리포트 생성
- [ ] 다중 정책 파일 지원
- [ ] 로깅 및 에러 핸들링 강화

## 문제 해결

### DRM 파일을 열 수 없는 경우
- Excel이 설치되어 있고 정상 작동하는지 확인
- xlwings가 Excel과 제대로 연결되어 있는지 확인: `xlwings.App().visible = True`로 테스트

### 헤더를 찾을 수 없는 경우
- 파일에 'Rulename'과 'Enable' 컬럼이 있는지 확인
- 컬럼 이름의 대소문자와 철자를 확인

### 성능이 여전히 느린 경우
- 파일 크기 확인 (1000행 이상인 경우 제한 확인)
- Excel 파일이 다른 프로그램에서 열려있는지 확인

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.
