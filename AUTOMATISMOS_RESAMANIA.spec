# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\logodeveloper.png', '.'), ('C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\feliz_cumpleanos.png', '.'), ('C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\PAGADEUDA.png', '.'), ('C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\PAGOHECHO.png', '.'), ('C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\config.json', '.')],
    hiddenimports=['win32com', 'win32com.client', 'pythoncom', 'pywintypes'],
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
    name='AUTOMATISMOS_RESAMANIA',
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
    icon=['C:\\Users\\JESUS\\IdeaProjects\\AUTOMASMOS RESAMANIA V1\\favicon.ico'],
)
