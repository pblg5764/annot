# -*- mode: python ; coding: utf-8 -*-
# AnotAI - Powered by Maehan Solutions
# Build with:  pyinstaller --clean --noconfirm AnotAI.spec

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],                      # brand assets are base64-embedded in brand.py
    hiddenimports=[
        'brand',
        'fitz',
        'pymupdf',
        'docx',
        'openpyxl',
        'openpyxl.cell._writer',
        'PIL._tkinter_finder',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'scipy', 'pandas',
              'PyQt6', 'PySide2', 'PySide6', 'IPython', 'jupyter'],
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
    name='AnotAI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/anotai.ico',
    version='version_info.txt',
)
