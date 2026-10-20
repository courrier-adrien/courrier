# -*- mode: python ; coding: utf-8 -*-
# Exe autonome anonymise (dossier "onedir") : tout est embarque, aucune donnee perso.
# Construire avec : pyinstaller COURRIER.spec --noconfirm --clean

a = Analysis(
    ['COURRIER.py'],
    pathex=['.'],
    binaries=[],
    datas=[('ressources', 'ressources')],
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
    name='COURRIER ADRIEN',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    icon='ressources\\ICONES\\courrier.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='COURRIER',
)
