# -*- mode: python ; coding: utf-8 -*-
# 汎用spec: version.py の __version__ を読み取り、exe名に自動反映
# 使用: pyinstaller DEF_Command_Set_App.spec --clean --noconfirm

import os
import re

_spec_dir = os.path.dirname(os.path.abspath(SPEC))
with open(os.path.join(_spec_dir, 'version.py'), encoding='utf-8') as _f:
    _match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', _f.read())
_version = _match.group(1) if _match else '0.00'
_app_name = f'DEF_Command_Set_App_v{_version}'


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['matplotlib.backends.backend_tkagg', 'openpyxl', 'win32com.client', 'pythoncom', 'PIL', 'PIL.Image', 'PIL.ImageGrab', 'PIL.ImageTk'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'tensorflow', 'keras', 'IPython', 'jedi', 'sympy', 'scipy', 'numba', 'llvmlite', 'lxml', 'cryptography', 'pygments'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=_app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
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
    name=_app_name,
)
