# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

import datetime
now = datetime.datetime.now()
VERSIONE = f"0.{now.year % 10}.{now.month}.{now.day}"

import glob
py_files = glob.glob("*.py")

a = Analysis(
    py_files,
    pathex=[],
    binaries=[],
    datas=[
        ('icon.ico', '.'),
        ('i18n/*', 'i18n')  
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=f'DesktopIconBackupManager_{VERSIONE}.exe',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    icon='icon.ico',
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)