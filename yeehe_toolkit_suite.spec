# -*- mode: python ; coding: utf-8 -*-

hiddenimports = [
    "uvicorn",
    "uvicorn.logging",
    "uvicorn.loops.auto",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets.auto",
    "fastapi",
    "starlette",
    "pydantic",
    "pypdf",
]
excludes = [
    "IPython",
    "jedi",
    "matplotlib",
    "pygame",
    "pytest",
    "scipy",
    "sklearn",
]

datas = [
    ("settings.json", "."),
    ("logo.png", "."),
]


a = Analysis(
    ["webui_launcher.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Yeehe_Toolkit_Suite",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    name="Yeehe_Toolkit_Suite",
)
