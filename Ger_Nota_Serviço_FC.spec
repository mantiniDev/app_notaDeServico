# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gerador_documentos.py'],
    pathex=[],
    binaries=[],
    datas=[('RobotoMono-Regular.ttf', '.'), ('RobotoMono-Bold.ttf', '.'), ('RobotoMono-Italic.ttf', '.'), ('icone.ico', '.')],
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
    name='Ger_Nota_Serviço_FC',
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
    icon=['icone.ico'],
)
