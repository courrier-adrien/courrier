# -*- mode: python ; coding: utf-8 -*-
# Exe autonome anonymise (dossier "onedir") : tout est embarque, aucune donnee perso.
# Construire avec : pyinstaller COURRIER.spec --noconfirm --clean
#
# PDF (22/11/2026) : pypdfium2 (Apache 2.0 / BSD) remplace PyMuPDF, sous AGPL,
# incompatible avec la PolyForm Noncommercial. PyMuPDF reste installe dans
# venv_build pour generer_pdf.py (outil de build, jamais distribue) : il est
# EXCLU ici pour ne jamais finir dans l'exe. copy_metadata embarque les
# licences de pypdfium2 et de PDFium (obligation de la licence Apache).
from PyInstaller.utils.hooks import copy_metadata

a = Analysis(
    ['COURRIER.py'],
    pathex=['.'],
    binaries=[],
    datas=[('ressources', 'ressources')] + copy_metadata('pypdfium2'),
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['fitz', 'pymupdf'],
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
