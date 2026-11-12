# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = collect_data_files('docx')
binaries = collect_dynamic_libs('pymupdf')

a = Analysis(
    ['statement_finder.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'fitz', 'pdfplumber', 'pypdf', 'pdfminer', 'pypdfium2',
        'cryptography', 'PIL', 'numpy', 'pandas', 'PyQt6', 'PySide6',
        'cv2', 'openpyxl', 'reportlab', 'pytest', 'hypothesis',
        'paramiko', 'uvicorn', 'fastapi',
        'lxml.html', 'setuptools', 'pkg_resources', 'win32com',
    ],
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
    name='Statement Finder',
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
)

