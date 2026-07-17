# -*- mode: python ; coding: utf-8 -*-
# Spec de PyInstaller para "Automatizador de Cartas".
# Generar el ejecutable con:  pyinstaller build.spec
#
# Nota sobre pywebview: en Windows usa el motor Edge WebView2 (viene
# preinstalado en Windows 10/11 modernos); en Linux necesita GTK o QT
# WebEngine instalado en el sistema; en macOS usa WebKit nativo.
# Compila SIEMPRE en el mismo sistema operativo donde se va a ejecutar
# el programa (PyInstaller no genera ejecutables multiplataforma).

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hidden_imports = (
    collect_submodules("docx")
    + collect_submodules("openpyxl")
    + collect_submodules("webview")
    + ["win32com", "win32com.client", "pythoncom"]
)

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[("web", "web")],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch', 'torchvision', 'tensorflow', 'matplotlib', 'scipy', 'numpy', 'sympy', 
        'streamlit', 'supabase', 'ultralytics', 'cv2', 'pandas', 'jedi', 
        'IPython', 'notebook', 'jinja2', 'tornado', 'bokeh', 'altair', 
        'pydeck', 'watchdog', 'fastapi', 'uvicorn', 'scikit-learn', 'sklearn'
    ],
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
    name="AutomatizadorCartas",
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
    icon=None,
)
