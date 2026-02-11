# 방화벽 정책 검증 자동화 스크립트

이 프로젝트는 방화벽 정책 변경이 성공적으로 적용되었는지 자동으로 검증하는 스크립트입니다.

## 개요

방화벽 정책 파일은 DRM(Digital Rights Management)이 적용되어 있어 일반적인 Excel 라이브러리로는 읽을 수 없습니다. 이 스크립트는 `xlwings`를 사용하여 DRM 보호 파일을 처리하고, 정책 변경 사항을 검증합니다.

## 주요 기능

1. **TUI 기반 파일 선택**: Rich 라이브러리를 사용한 사용자 친화적인 인터페이스
   - 현재 디렉터리에서 Excel 파일 자동 탐지
   - 여러 대상 정책 파일 선택 지원

2. **정책 파일 파싱**: DRM이 적용된 Excel 파일에서 정책 정보 추출
   - `Rulename`: 정책 이름
   - `Enable`: 정책 활성화 상태 (Y/N 형식 지원)

3. **성능 최적화**: 전체 범위를 한 번에 읽어 처리 속도 향상

4. **빈 셀 처리**: 정책 파일의 많은 빈 공간을 자동으로 건너뛰고 실제 데이터만 추출

5. **고급 정책 검증**:
   - 대상 정책의 삭제/비활성화 확인
   - 대상 외에 삭제되거나 비활성화된 정책 탐지
   - 덜 삭제/비활성화된 정책 탐지 (대상에는 있지만 실제로는 처리 안됨)

6. **리포트 조회**: 검증 결과를 TUI로 조회 및 재조회 가능

## 워크플로우

```
┌─────────────────────────────────────────────────────────────┐
│ 1. 정책 파일 준비                                            │
│    - 현재 디렉터리에 Excel 파일 배치                          │
│    - running_policy.xlsx (현재 실행 중인 정책)              │
│    - candidate_policy.xlsx (변경 예정 정책)                 │
│    - target_policies.xlsx (대상 정책, 여러 개 가능)         │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 2. TUI로 파일 선택                                           │
│    - 현재 디렉터리에서 Excel 파일 자동 탐지                  │
│    - Running/Candidate 정책 파일 선택                      │
│    - 대상 정책 파일 선택 (여러 개 선택 가능)                │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 3. 정책 파일 파싱                                            │
│    - xlwings로 DRM 보호 파일 열기                           │
│    - 전체 데이터를 한 번에 읽기 (성능 최적화)                │
│    - 헤더 행 자동 탐지 ('Rulename', 'Enable')               │
│    - 빈 셀 제거 및 데이터 정제                               │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 4. 대상 정책 목록 로드                                       │
│    - 대상 정책 파일에서 작업구분이 "삭제"인 정책 추출        │
│    - "Rule Name", "Rulename", "Policy Name" 컬럼 지원      │
│    - 여러 파일에서 정책 목록 통합                            │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 5. 정책 검증                                                 │
│    - Running vs Candidate 비교                              │
│    - 대상 정책의 삭제/비활성화 확인                          │
│    - 대상 외 정책의 삭제/비활성화 탐지                      │
│    - 덜 삭제/비활성화된 정책 탐지                            │
└─────────────────────────────────────────────────────────────┘
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ 6. 검증 결과 리포트 생성 및 조회                             │
│    - Excel 리포트 생성 (validation_report.xlsx)             │
│    - TUI로 리포트 조회 및 재조회                             │
└─────────────────────────────────────────────────────────────┘
```

## 사용 방법

### 기본 사용 (전체 워크플로우 실행)

```bash
python parse_firewall_policy.py
```

스크립트를 실행하면 TUI 인터페이스가 나타나며 다음 단계를 진행합니다:

1. **Running 정책 파일 선택**: 현재 디렉터리의 Excel 파일 목록에서 선택
2. **Candidate 정책 파일 선택**: 현재 디렉터리의 Excel 파일 목록에서 선택
3. **대상 정책 파일 선택**: 여러 개 선택 가능 (쉼표로 구분, 예: 1,2,3)
4. **정책 파일 파싱**: 선택된 파일들을 자동으로 파싱
5. **대상 정책 목록 로드**: 작업구분이 "삭제"인 정책 추출
6. **정책 검증**: 변경 사항 검증 수행
7. **검증 결과 리포트 저장**: `validation_report.xlsx` 생성
8. **리포트 조회**: 검증 결과를 TUI로 조회

### 필요한 파일

- **Running 정책 파일**: 현재 실행 중인 정책 파일 (DRM 보호 가능)
- **Candidate 정책 파일**: 변경 예정 정책 파일 (DRM 보호 가능)
- **대상 정책 파일(들)**: 작업구분 컬럼이 "삭제"인 정책이 포함된 파일 (여러 개 선택 가능, 선택사항)

모든 파일은 **현재 디렉터리**에 있어야 하며, 스크립트 실행 시 TUI로 선택합니다.

### 프로그래밍 방식 사용

```python
from parse_firewall_policy import (
    parse_policy_file, 
    parse_target_file, 
    load_target_policies, 
    validate_policy_changes
)

# 1. 정책 파일 파싱 (Rulename, Enable 컬럼 추출)
running_df = parse_policy_file("running_policy.xlsx")
candidate_df = parse_policy_file("candidate_policy.xlsx")

# 2. 대상 정책 파일 파싱 (작업구분이 "삭제"인 행만 추출)
# "Rule Name", "Rulename", "Policy Name" 컬럼 모두 지원
target_policies = parse_target_file("target_policies.xlsx")

# 또는 간단한 텍스트/Excel 파일에서 로드
# target_policies = load_target_policies("target_policies.txt")

# 3. 정책 변경 사항 검증
validation_results = validate_policy_changes(
    running_df, 
    candidate_df, 
    target_policies
)

# 4. 결과 출력 및 저장
print(validation_results)
validation_results.to_excel("validation_report.xlsx", index=False)

# 검증 결과 요약
deleted_count = len(validation_results[validation_results['Status'] == 'DELETED'])
disabled_count = len(validation_results[validation_results['Status'] == 'DISABLED'])
print(f"\n삭제 확인: {deleted_count}개")
print(f"비활성화 확인: {disabled_count}개")
```

### 파일 구조

```
validate-policy/
├── parse_firewall_policy.py      # 메인 스크립트
├── requirements.txt              # 필요한 패키지 목록
├── running_policy.xlsx            # 현재 실행 중인 정책 파일 (선택)
├── candidate_policy.xlsx          # 변경 예정 정책 파일 (선택)
├── target_policies.xlsx          # 대상 정책 파일 (작업구분 컬럼 포함, 선택사항)
├── validation_report.xlsx         # 검증 결과 리포트 (자동 생성)
└── README.md                     # 이 파일
```

**참고**: 처리 중간 파일(`*_processed.xlsx`)은 더 이상 생성되지 않습니다.

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

### `parse_target_file(file_path: str) -> list`

대상 정책 파일을 파싱하여 정책 이름 목록을 추출합니다.

**주요 기능:**
- "Rule Name", "Rulename", "Policy Name" 컬럼 모두 지원 (대소문자 무시)
- "작업구분" (Task Type) 컬럼이 있으면 값이 "삭제" (Delete)인 행만 추출
- Enable 컬럼은 없음 (정책 이름만 추출)
- DRM 보호 파일 지원

**파라미터:**
- `file_path`: 대상 정책 파일 경로 (Excel 파일, DRM 보호 가능)

**반환값:**
- `list`: 정책 이름 리스트 (중복 제거됨)

**예제:**
```python
# 대상 파일에서 작업구분이 "삭제"인 정책만 추출
target_policies = parse_target_file("target_policies.xlsx")
```

### `load_target_policies(file_path: str) -> list`

대상 정책 이름 목록을 파일에서 읽어옵니다. (간단한 텍스트/Excel 파일용)

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
openpyxl>=3.0.0
rich>=13.0.0
```

설치:
```bash
pip install -r requirements.txt
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

## 검증 결과 상태

검증 리포트의 `Status` 컬럼은 다음 값을 가질 수 있습니다:

### 대상 정책 검증 결과

- `DELETED` ✓: 정책이 삭제되었습니다 (Running에는 있지만 Candidate에는 없음)
- `DISABLED` ✓: 정책이 비활성화되었습니다 (Y → N)
- `NOT_DISABLED` ⚠: 정책이 비활성화되지 않았습니다 (대상에는 있지만 Y 상태 유지)
- `RE_ENABLED` ⚠: 정책이 다시 활성화되었습니다 (N → Y)
- `NO_CHANGE`: 변경 없음
- `NOT_IN_RUNNING`: Running 정책에 존재하지 않음
- `CHANGED`: Enable 상태가 변경됨

### 대상 외 정책 검증 결과

- `UNEXPECTED_DELETED` ⚠: 대상 외 정책이 삭제되었습니다
- `UNEXPECTED_DISABLED` ⚠: 대상 외 정책이 비활성화되었습니다 (Y → N)

**Enable 값**: Y (활성화) 또는 N (비활성화)

## 주요 개선 사항

- [x] TUI 기반 파일 선택 (Rich 라이브러리)
- [x] 여러 대상 정책 파일 선택 지원
- [x] 처리 중간 파일 저장 제거
- [x] Enable 값 Y/N 형식 지원
- [x] 대상 외 정책 검증 기능 추가
- [x] 덜 삭제/비활성화된 정책 탐지
- [x] 리포트 조회 기능 (TUI)

## 향후 개발 계획

- [ ] HTML 리포트 생성
- [ ] 이메일 알림 기능
- [ ] 로깅 및 에러 핸들링 강화
- [ ] 설정 파일 지원

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
