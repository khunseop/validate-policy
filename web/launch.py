"""
exe 진입점: import/실행 예외를 잡아 오류 로그를 남깁니다.
실행: python web/launch.py (프로젝트 루트에서) 또는 빌드 후 validate-policy-web.exe
"""
import sys
import os
from pathlib import Path

# 비빌드 시 web 디렉터리를 path에 추가 (from app import app 용)
if not getattr(sys, 'frozen', False):
    _web_dir = Path(__file__).resolve().parent
    sys.path.insert(0, str(_web_dir))

def _log_error():
    import traceback
    tb = traceback.format_exc()
    err_file = os.path.join(os.environ.get('TEMP', os.path.expanduser('~')), 'validate_policy_web_error.txt')
    try:
        with open(err_file, 'w', encoding='utf-8') as f:
            f.write(tb)
        print(tb, file=sys.stderr)
        print('오류 내용이 저장됨:', err_file, file=sys.stderr)
    except Exception:
        print(tb, file=sys.stderr)
    if getattr(sys, 'frozen', False):
        input('Enter 키로 종료...')

def main():
    try:
        from app import app
        app.run(host='127.0.0.1', port=5000, debug=False)
    except Exception:
        _log_error()
        sys.exit(1)

if __name__ == '__main__':
    main()
