# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec: Flask 웹 앱 (방화벽 정책 검증)
# 빌드: pyinstaller web.spec  (프로젝트 루트에서 실행)

from pathlib import Path

project_root = Path(SPECPATH)
web_dir = project_root / 'web'

# Flask 템플릿 번들에 포함 (실행 시 sys._MEIPASS 아래로 풀림)
datas = [
    (str(web_dir / 'templates'), 'templates'),
]
# static 폴더가 있으면 추가
if (web_dir / 'static').exists():
    datas.append((str(web_dir / 'static'), 'static'))

a = Analysis(
    scripts=[str(web_dir / 'launch.py')],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'flask',
        'werkzeug',
        'jinja2',
        'core',
        'core.parser',
        'core.validator',
        'core.utils',
        'core.vendor',
        'core.__init__',
        'pandas',
        'openpyxl',
        'xlwings',
        'rich',
        'rich.console',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='validate-policy-web',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
