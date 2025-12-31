# -*- mode: python ; coding: utf-8 -*-
import os
import glob

block_cipher = None

# --- VERSIONING (Read from version.txt) ---
try:
    with open("version.txt", "r", encoding="utf-8") as f:
        VERSIONE = f.read().strip()
except Exception:
    VERSIONE = "0.0.0"
# ------------------------------------------

py_files = glob.glob("*.py")

a = Analysis(
    py_files,
    pathex=[],
    binaries=[],
    datas=[
        ('icon.ico', '.'),
        ('i18n/*', 'i18n'),
        ('version.txt', '.')
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
    name=f'desktop-icon-backup-manager_{VERSIONE}.exe',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)