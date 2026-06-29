# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, collect_submodules

hiddenimports = (
    collect_submodules("win32com")
    + collect_submodules("pypdf")
    + collect_submodules("pdfplumber")
    + collect_submodules("PIL")
)
datas = [
    ("config.json", "."),
    ("office_bridge.ps1", "."),
]
binaries = []

# RapidOCR 的模型和 ONNX Runtime/OpenCV DLL 必须显式收集，确保 onefile
# 解压后可以从 PyInstaller 的临时目录加载模型和本地运行库。
for package in ("rapidocr_onnxruntime", "onnxruntime", "cv2", "numpy", "PIL"):
    package_datas, package_binaries, package_hidden = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hidden

a = Analysis(
    ["gui_app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="月度工作量表自动生成器",
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
