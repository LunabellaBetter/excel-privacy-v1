# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


PROJECT_ROOT = Path(SPECPATH)

# The integrity verifier hashes these exact source/resource files at runtime.
# PyInstaller normally embeds Python modules only in its archive, so explicit
# data copies are required at the paths expected by excel_privacy_tool.py.
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
    ['privacy_app.py'],
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
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Excel Privacy Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch="arm64",
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
    name='Excel Privacy Tool',
)
app = BUNDLE(
    coll,
    name='Excel Privacy Tool.app',
    version='1.0.0',
    info_plist={'CFBundleShortVersionString': '1.0.0', 'CFBundleVersion': '1.0.0', 'LSMinimumSystemVersion': '11.0', 'NSHumanReadableCopyright': 'Copyright (c) 2026 LunaBella'},
    icon=None,
    bundle_identifier="io.lunabella.excelprivacy",
)
