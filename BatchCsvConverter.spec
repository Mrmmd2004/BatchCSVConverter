# -*- mode: python ; coding: utf-8 -*-
# Build:  python -m PyInstaller --noconfirm --clean BatchCsvConverter.spec
# Output: dist\BatchCsvConverter.exe  (single-file, windowed, x64 when built with 64-bit Python)

a = Analysis(
    ['batch_csv_converter.py'],
    pathex=[],
    binaries=[],
    datas=[('clubapp.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'pandas', 'matplotlib', 'PIL', 'scipy', 'IPython', 'pytest'],
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
    name='BatchCsvConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,              # UPX often triggers antivirus false positives
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['clubapp.ico'],
    version='version_info.txt',
)
