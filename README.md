# 방화벽 정책 검증 자동화 시스템

이 프로젝트는 방화벽 정책 변경이 성공적으로 적용되었는지 자동으로 검증하는 시스템입니다.

## 프로젝트 구조

```
validate-policy/
├── core/                    # 핵심 로직 모듈
│   ├── __init__.py
│   ├── parser.py           # 파일 파싱 로직
│   ├── validator.py        # 검증 로직
│   └── utils.py            # 유틸리티 함수
├── cli/                     # CLI 인터페이스
│   ├── __init__.py
│   └── main.py            # CLI 메인 진입점
├── web/                     # 웹 인터페이스
│   ├── app.py              # Flask 웹 애플리케이션
│   ├── templates/          # HTML 템플릿
│   └── static/            # 정적 파일
├── parse_firewall_policy.py # CLI 진입점 (기존 호환성)
├── requirements.txt        # 필요한 패키지
└── README.md              # 이 파일
```

## 주요 기능

1. **정책 파일 파싱**: Excel 파일에서 정책 정보 추출
   - `Rulename`: 정책 이름
   - `Enable`: 정책 활성화 상태 (Y/N)

2. **대상 정책 파일 파싱**:
   - "Rule Name", "Rulename", "Policy Name" 컬럼 지원
   - "작업구분" 컬럼이 있으면 "삭제"인 행만 추출 (없으면 모든 행)
   - "제외사유" 컬럼이 있으면 빈 데이터인 행만 추출

3. **정책 검증**:
   - 대상 정책의 삭제/비활성화 확인
   - 대상 외 정책의 삭제/비활성화 탐지
   - 덜 삭제/비활성화된 정책 탐지

4. **성능 최적화**: 딕셔너리 기반 O(1) 조회로 빠른 처리

## 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### CLI 버전 (터미널)

```bash
python parse_firewall_policy.py
```

또는

```bash
python -m cli.main
```

### 웹 버전 (브라우저)

```bash
python web/app.py
```

브라우저에서 `http://127.0.0.1:5000` 접속

## 모듈 설명

### core/parser.py
- `parse_policy_file()`: 정책 파일 파싱 (Rulename, Enable 추출)
- `parse_target_file()`: 대상 정책 파일 파싱

### core/validator.py
- `validate_policy_changes()`: 정책 변경 사항 검증
- `normalize_enable()`: Enable 값 정규화 (Y/N)

### core/utils.py
- `show_summary()`: CLI용 결과 요약 표시
- `get_summary_dict()`: 웹용 결과 요약 딕셔너리 생성

### cli/main.py
- CLI 인터페이스 메인 함수
- TUI로 파일 선택 및 검증 진행

### web/app.py
- Flask 웹 애플리케이션
- 파일 업로드 및 검증 API

## 검증 결과 상태

- `DELETED` ✓: 정책이 삭제되었습니다
- `DISABLED` ✓: 정책이 비활성화되었습니다 (Y → N)
- `NOT_DISABLED` ⚠: 정책이 비활성화되지 않았습니다
- `UNEXPECTED_DELETED` ⚠: 대상 외 정책이 삭제되었습니다
- `UNEXPECTED_DISABLED` ⚠: 대상 외 정책이 비활성화되었습니다
- `RE_ENABLED` ⚠: 정책이 다시 활성화되었습니다
- `NO_CHANGE`: 변경 없음
- `NOT_IN_RUNNING`: Running 정책에 존재하지 않음

## 개발자 가이드

### 새로운 기능 추가

1. 핵심 로직은 `core/` 모듈에 추가
2. CLI 기능은 `cli/main.py`에 추가
3. 웹 기능은 `web/app.py`에 추가

### 모듈 간 의존성

- `core/`: 독립적 (다른 모듈에 의존 없음)
- `cli/`: `core/` 의존
- `web/`: `core/` 의존

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.
