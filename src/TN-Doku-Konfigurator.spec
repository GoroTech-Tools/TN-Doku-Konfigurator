# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from pathlib import Path

ROOT_DIR = Path(os.getcwd()).resolve()
SRC_DIR = ROOT_DIR / 'src'

# Build-Informationen laden
sys.path.insert(0, str(SRC_DIR))
from build_info import BUILD_INFO

version = BUILD_INFO['version']

a = Analysis(
    [str(SRC_DIR / 'main.py')],
    pathex=[str(SRC_DIR)],
    binaries=[],
    datas=[
        (str(SRC_DIR / 'build_info.py'), '.'),
        (str(SRC_DIR / 'app_icon.ico'), '.'),
    ],
    hiddenimports=[],
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
    exclude_binaries=False,
    name='TN-Doku-Konfigurator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(SRC_DIR / 'app_icon.ico'),
)
