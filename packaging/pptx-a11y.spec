# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None
spec_dir = Path(SPECPATH).resolve()
project_root = spec_dir.parent
src_path = project_root / "src"

datas = [
    (str(src_path / "pptx_a11y" / "reports" / "themes"), "pptx_a11y/reports/themes"),
]
datas += collect_data_files("pptx")
datas += collect_data_files("fitz")

hiddenimports = [
    "PySide6",
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "pptx",
    "pptx_a11y",
    "pptx_a11y.gui",
    "pptx_a11y.gui.app",
    "pptx_a11y.gui.main_window",
    "pptx_a11y.gui.models",
    "pptx_a11y.gui.theme",
    "pptx_a11y.gui.triage_dialog",
    "pptx_a11y.gui.worker",
    "fitz",
    "pikepdf",
]
hiddenimports += collect_submodules("pptx_a11y")

a = Analysis(
    [str(src_path / "pptx_a11y" / "gui" / "app.py")],
    pathex=[str(src_path)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="pptx-a11y",
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="pptx-a11y",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="pptx-a11y.app",
        icon=None,
        bundle_identifier="com.thirstyhead.pptx-a11y",
        info_plist={
            "CFBundleDisplayName": "pptx-a11y",
            "CFBundleName": "pptx-a11y",
            "CFBundleIdentifier": "com.thirstyhead.pptx-a11y",
            "CFBundleVersion": "0.3.0",
            "CFBundleShortVersionString": "0.3.0",
            "NSHighResolutionCapable": "True",
            "LSApplicationCategoryType": "public.app-category.utilities",
        },
    )
