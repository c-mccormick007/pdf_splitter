# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pdf_splitter.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\chris.mccormick\\Desktop\\AP Automation\\pdf_splitter\\.venv\\Lib\\site-packages\\tkinterdnd2\\tkdnd\\win-x64\\', 'libtkdnd2.9.4')],
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
    [],
    exclude_binaries=True,
    name='pdf_splitter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pdf_splitter',
)
