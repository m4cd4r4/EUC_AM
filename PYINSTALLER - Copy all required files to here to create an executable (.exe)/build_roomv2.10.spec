# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['build_roomv2.10.py'],
    pathex=[],
    binaries=[],
    datas=[('EUC_Perth_Assets.xlsx', '.'), ('Plots', 'Plots'), ('inventory-levels_4.2v1.py', '.'), ('inventory-levels_BRv1.py', '.'), ('inventory-levels_combinedv1.1.py', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='build_roomv2.10',
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
