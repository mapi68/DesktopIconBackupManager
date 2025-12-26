# -*- mode: python ; coding: utf-8 -*-
import os
import datetime
import glob

block_cipher = None

# --- AUTOMATIC VERSIONING (Based on source file modification) ---
# This ensures the version only changes when the code is actually edited
try:
    script_path = "DesktopIconBackupManager.py"
    if os.path.exists(script_path):
        mtime = os.path.getmtime(script_path)
        last_mod = datetime.datetime.fromtimestamp(mtime)
        VERSIONE = f"0.{last_mod.year % 10}.{last_mod.month}.{last_mod.day}"
    else:
        VERSIONE = "0.0.0"
except Exception:
    VERSIONE = "0.0.0"
# ----------------------------------------------------------------

# Automatically include all .py files in the directory
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
    # The output filename now matches the source code version
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