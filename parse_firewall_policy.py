"""
방화벽 정책 검증 CLI 진입점

이 스크립트는 CLI 인터페이스를 제공합니다.
웹 버전을 사용하려면 web/app.py를 실행하세요.
"""

import sys
from pathlib import Path

# CLI 모듈 경로 추가
sys.path.insert(0, str(Path(__file__).parent))

from cli.main import main

if __name__ == "__main__":
    main()
