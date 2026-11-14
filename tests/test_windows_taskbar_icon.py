"""Specialized test suite for Windows taskbar icon, AUMID, dialog inheritance and fallbacks."""

import ast
from pathlib import Path
import struct
import sys
from unittest.mock import MagicMock, patch

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import privacy_app
from privacy_handler.excel_privacy_tool import LOCKED_FILES, verify_self_before_sensitive


def test_windows_aumid_setting_on_win32():
    """Verify that on Windows, SetCurrentProcessExplicitAppUserModelID is called with the stable ID."""
    mock_shell32 = MagicMock()
    mock_ctypes = MagicMock()
    mock_ctypes.windll.shell32 = mock_shell32

    with patch.object(sys, "platform", "win32"), patch.dict("sys.modules", {"ctypes": mock_ctypes}):
        privacy_app._set_windows_aumid()
        mock_shell32.SetCurrentProcessExplicitAppUserModelID.assert_called_once_with(
            "io.lunabella.excelprivacy"
        )


def test_windows_aumid_non_win32_skips():
    """Verify that on macOS/Linux, Windows ctypes APIs are never invoked."""
    mock_ctypes = MagicMock()
    with patch.object(sys, "platform", "darwin"), patch.dict("sys.modules", {"ctypes": mock_ctypes}):
        privacy_app._set_windows_aumid()
        mock_ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID.assert_not_called()


def test_windows_icon_non_win32_skips():
    """Verify that on macOS/Linux, _apply_windows_icon exits immediately without calling iconbitmap."""
    app_mock = MagicMock()
    with patch.object(sys, "platform", "darwin"):
        privacy_app.PrivacyApp._apply_windows_icon(app_mock)
        app_mock.iconbitmap.assert_not_called()


def test_windows_icon_candidate_dev_resolution(tmp_path):
    """Verify resolution of development path assets/brand/icons/win/ExcelPrivacy.ico."""
    app_mock = MagicMock()
    fake_ico = ROOT / "assets" / "brand" / "icons" / "win" / "ExcelPrivacy.ico"

    with patch.object(sys, "platform", "win32"), \
         patch.object(sys, "frozen", False, create=True), \
         patch.object(privacy_app, "_set_windows_aumid") as mock_aumid:
        privacy_app.PrivacyApp._apply_windows_icon(app_mock)
        mock_aumid.assert_called_once()
        app_mock.iconbitmap.assert_called_once_with(default=str(fake_ico))


def test_windows_icon_candidate_meipass_resolution(tmp_path):
    """Verify resolution of PyInstaller sys._MEIPASS path."""
    app_mock = MagicMock()
    meipass_ico = tmp_path / "assets" / "brand" / "icons" / "win" / "ExcelPrivacy.ico"
    meipass_ico.parent.mkdir(parents=True, exist_ok=True)
    meipass_ico.write_bytes(b"\x00\x00\x01\x00")

    with patch.object(sys, "platform", "win32"), \
         patch.object(sys, "_MEIPASS", str(tmp_path), create=True), \
         patch.object(privacy_app, "_set_windows_aumid"):
        privacy_app.PrivacyApp._apply_windows_icon(app_mock)
        app_mock.iconbitmap.assert_called_once_with(default=str(meipass_ico))


def test_windows_icon_fallback_when_ico_missing(tmp_path):
    """Verify fallback to sys.executable when all ICO candidates are missing in frozen mode."""
    app_mock = MagicMock()
    fake_exe = (tmp_path / "Excel Privacy Tool.exe").resolve()
    fake_exe.write_bytes(b"MZ\x90\x00")

    def mock_is_file(path_obj):
        return path_obj.resolve() == fake_exe

    with patch.object(sys, "platform", "win32"), \
         patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)), \
         patch.object(Path, "is_file", autospec=True, side_effect=mock_is_file), \
         patch.object(privacy_app, "_set_windows_aumid"):
        privacy_app.PrivacyApp._apply_windows_icon(app_mock)
        app_mock.iconbitmap.assert_called_once_with(default=str(fake_exe))


def test_dialogs_inherit_from_toplevel():
    """Verify all modal dialogs are Toplevel widgets inheriting default icon on Windows."""
    import tkinter as tk

    assert issubclass(privacy_app.RecordKeyDialog, tk.Toplevel)
    assert issubclass(privacy_app.SharedGroupDialog, tk.Toplevel)
    assert issubclass(privacy_app.PromptPreviewDialog, tk.Toplevel)


def test_windows_spec_configuration():
    """Verify Excel Privacy Tool Windows.spec datas, locked resources and icons."""
    spec_path = ROOT / "Excel Privacy Tool Windows.spec"
    content = spec_path.read_text(encoding="utf-8")
    assert 'datas.append((str(PROJECT_ROOT / "assets" / "brand" / "icons" / "win" / "ExcelPrivacy.ico"), "assets/brand/icons/win"))' in content
    assert 'icon=str(PROJECT_ROOT / "assets" / "brand" / "icons" / "win" / "ExcelPrivacy.ico")' in content
    assert 'MICROSOFT_RUNTIME_NAMES = {"vcruntime140.dll", "vcruntime140_1.dll", "ucrtbase.dll"}' in content


def test_excel_privacy_ico_integrity():
    """Verify that ExcelPrivacy.ico exists and has 7 valid multi-resolution icon entries."""
    ico_path = ROOT / "assets" / "brand" / "icons" / "win" / "ExcelPrivacy.ico"
    assert ico_path.is_file(), "ExcelPrivacy.ico must exist"
    data = ico_path.read_bytes()
    assert len(data) > 0

    reserved, type_, count = struct.unpack("<HHH", data[:6])
    assert reserved == 0 and type_ == 1, "Must be valid ICO header"
    assert count == 7, f"Expected 7 icon entries, got {count}"

    sizes = []
    for i in range(count):
        offset = 6 + i * 16
        w, h = struct.unpack("<BB", data[offset:offset + 2])
        actual_w = 256 if w == 0 else w
        actual_h = 256 if h == 0 else h
        sizes.append((actual_w, actual_h))

    expected_sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    assert sizes == expected_sizes


def test_verify_self_integrity_passes():
    """Verify that verify_self_before_sensitive passes with current approved_checksums."""
    verify_self_before_sensitive()
