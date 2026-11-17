# -*- mode: python ; coding: utf-8 -*-
# В .spec недоступен __file__; PyInstaller определяет SPECPATH — путь к папке со spec.
import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

ROOT = SPECPATH  # noqa: F821

datas = []
for icon_rel in ('icons/app.ico', 'icons/window.png', 'icons/window.ico'):
    icon_abs = os.path.join(ROOT, icon_rel)
    if os.path.exists(icon_abs):
        datas.append((icon_abs, 'icons'))

datas += collect_data_files('docx')
hiddenimports = collect_submodules('docx')

_app_ico = os.path.join(ROOT, 'icons', 'app.ico')
exe_icon = _app_ico if os.path.exists(_app_ico) else None

main_py = os.path.join(ROOT, 'main.py')

a = Analysis(
    [main_py],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CRF',
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
    icon=exe_icon,
)