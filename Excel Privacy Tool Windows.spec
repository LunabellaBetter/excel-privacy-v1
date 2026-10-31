# -*- mode: python ; coding: utf-8 -*-
"""Windows x64 one-folder build; run with the pinned PyInstaller on Windows."""

import sys
from pathlib import Path


if sys.platform != "win32":
    raise RuntimeError("Build this Windows application on Windows x64")
if sys.maxsize <= 2**32:
    raise RuntimeError("Use 64-bit Python for the Windows build")

PROJECT_ROOT = Path(SPECPATH)

# Keep this list aligned with excel_privacy_tool.LOCKED_FILES and the Mac spec.
# The verifier reads these files from the frozen resource directory at runtime.
LOCKED_RESOURCES = [
    "privacy_app.py",
    "privacy_handler/gui_controller.py",
    "privacy_handler/packaging_smoke.py",
    "privacy_handler/excel_privacy_tool.py",
    "privacy_handler/identity.py",
    "privacy_handler/key_detector.py",
    "privacy_handler/numeric_masking.py",
    "privacy_handler/conflict_engine.py",
    "privacy_handler/restore_engine.py",
    "privacy_handler/prompt_generator.py",
    "privacy_handler/table_scope.py",
    "privacy_handler/recovery_ledger.py",
    "privacy_handler/README.md",
    "privacy_handler/approved_checksums.json",
]

datas = [
    (str(PROJECT_ROOT / relative_path), str(Path(relative_path).parent))
    for relative_path in LOCKED_RESOURCES
]
datas.append((str(PROJECT_ROOT / "README.md"), "."))
datas.append((str(PROJECT_ROOT / "LICENSE"), "."))
datas.append((str(PROJECT_ROOT / "THIRD_PARTY_NOTICES.md"), "."))
datas.append((str(PROJECT_ROOT / "release" / "licenses"), "licenses"))


a = Analysis(
    ["privacy_app.py"],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
# Use Windows 10/11 system components and the user's official Microsoft VC++
# x64 runtime installation. This project does not redistribute Microsoft DLLs.
MICROSOFT_RUNTIME_NAMES = {"vcruntime140.dll", "vcruntime140_1.dll", "ucrtbase.dll"}
a.binaries = [
    item for item in a.binaries
    if Path(item[0]).name.lower() not in MICROSOFT_RUNTIME_NAMES
    and not Path(item[0]).name.lower().startswith("api-ms-win-")
]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Excel Privacy Tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="Excel Privacy Tool",
)
