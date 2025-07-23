# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['OnePage.py'],
    pathex=['C:\\scripts\\OnePage'],
    binaries=[],
    datas=[
        ('macros/Dashboard of Operations.xlsm', 'macros'),
        ('script.vbs', '.'),
        ('run.bat', '.'),
        ('assets/python.ico', 'assets')
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OnePage',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=_
