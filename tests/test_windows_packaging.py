"""Static gates for the Windows build configuration; no Windows runtime claim."""

import ast
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from privacy_handler.excel_privacy_tool import LOCKED_FILES


def _locked_resources(spec_path: Path) -> set[str]:
    tree = ast.parse(spec_path.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "LOCKED_RESOURCES"
            for target in node.targets
        ):
            return set(ast.literal_eval(node.value))
    raise AssertionError("Missing locked resources in PyInstaller spec")


def test_windows_spec_copies_exactly_the_locked_resources():
    windows_spec = ROOT / "Excel Privacy Tool Windows.spec"
    mac_spec = ROOT / "Excel Privacy Tool.spec"
    expected = {path.relative_to(ROOT).as_posix() for path in LOCKED_FILES}
    expected.add("privacy_handler/approved_checksums.json")

    assert _locked_resources(windows_spec) == expected
    assert _locked_resources(windows_spec) == _locked_resources(mac_spec)
    assert all((ROOT / path).is_file() for path in expected)


def test_windows_spec_is_one_folder_and_excludes_private_data():
    source = (ROOT / "Excel Privacy Tool Windows.spec").read_text(encoding="utf-8")
    tree = ast.parse(source)
    calls = {
        node.func.id
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
    }

    assert "COLLECT" in calls
    assert "BUNDLE" not in calls
    assert "privacy_app.py" in source
    assert "sys.platform != \"win32\"" in source
    assert "sys.maxsize <= 2**32" in source
    for private_path in (
        "private_keys/", "original_data/", "safe_for_codex/",
        "restored_outputs/", "archive/", "scratch/", "examples/",
    ):
        assert private_path not in source
