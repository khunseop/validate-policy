# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec: CLI 앱 (방화벽 정책 검증)
# 빌드: pyinstaller cli.spec  (프로젝트 루트에서 실행)

from pathlib import Path

project_root = Path(SPECPATH)  # PyInstaller가 spec 위치로 설정

a = Analysis(
    scripts=[str(project_root / 'parse_firewall_policy.py')],
    pathex=[str(project_root)],
    binaries=[],
    datas=[],
    hiddenimports=[
        'cli',
        'cli.main',
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
        'rich.prompt',
        'rich.panel',
        'rich.table',
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
    name='validate-policy-cli',
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
