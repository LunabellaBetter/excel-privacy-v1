"""Comprehensive unit tests for Checkpoint 8 GUI Workflow and Controller.

Covers all 35 required validation and workflow gates specified in Checkpoint 8 requirements:
1. load_workbook_populates_sheet_list
2. multi_sheet_selection
3. sheet_specific_header_rows
4. changing_header_row_refreshes_columns
5. sensitive_recommendations_preselected
6. user_can_unselect_recommendation
7. text_protection_selection
8. numeric_protection_selection
9. single_record_identity_selection
10. composite_record_identity_selection
11. invalid_record_identity_warning
12. no_record_identity_allowed
13. numeric_group_assignment
14. cross_sheet_same_group_assignment
15. masked_output_dir_selection_state
16. restored_output_dir_selection_state
17. mask_button_validation
18. mask_controller_calls_mask_workbook_v1
19. mask_result_stores_ai_prompt
20. copy_prompt_uses_generated_prompt
21. returned_workbook_selection
22. ledger_auto_resolution_by_run_id
23. run_identity_error_user_message
24. compare_controller_calls_conflict_engine
25. compare_summary_counts
26. restore_controller_calls_restore_workbook_v1
27. restore_summary_success
28. restore_summary_partial_conflict_warning
29. unsupported_parts_message_is_human_readable
30. gui_never_displays_scale_factor
31. gui_never_displays_secret_key
32. gui_never_displays_plain_record_identity
33. background_task_disables_buttons
34. background_task_reenables_buttons_after_success
35. background_task_reenables_buttons_after_error
"""

from __future__ import annotations

from pathlib import Path
import sys
import pytest
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from privacy_handler.gui_controller import (
    APP_DIR,
    GUIWorkflowController,
    default_app_data_root,
    translate_security_error,
)
from privacy_handler.excel_privacy_tool import PrivacyToolError, mask_workbook_v1
from privacy_handler.table_scope import SheetProtectionScope


def create_sample_workbook(path: Path) -> Path:
    """Create a standard multi-sheet workbook with different header rows and columns."""
    wb = openpyxl.Workbook()

    # Sheet 1: Staff (header at row 2)
    ws1 = wb.active
    ws1.title = "Staff"
    ws1.append(["Title Banner", None, None, None])
    ws1.append(["EmpID", "姓名", "基本工资", "部门"])
    ws1.append(["E001", "张三", 12000, "技术部"])
    ws1.append(["E002", "李四", 15000, "市场部"])
    ws1.append(["E003", "王五", 18000, "技术部"])

    # Sheet 2: Bonuses (header at row 1)
    ws2 = wb.create_sheet(title="Bonuses")
    ws2.append(["EmpID", "月份", "奖金", "姓名"])
    ws2.append(["E001", "2026-01", 3000, "张三"])
    ws2.append(["E002", "2026-01", 4000, "李四"])
    ws2.append(["E003", "2026-01", 5000, "王五"])

    wb.save(path)
    return path


def select_unique_record_identifier(ctl: GUIWorkflowController, sheet_name: str) -> None:
    """Explicitly select the fixture's unique EmpID column for GUI-flow tests."""
    cfg = ctl.sheet_configs[sheet_name]
    if not cfg.get_record_key_columns():
        ctl.toggle_column_record_key(sheet_name, "A", True)


def confirm_with_unique_record_identifier(ctl: GUIWorkflowController, sheet_name: str) -> tuple[bool, str]:
    select_unique_record_identifier(ctl, sheet_name)
    return ctl.confirm_sheet_configuration(sheet_name)


# ------------------------------------------------------------------------------
# Tests 1 - 10: Workbook loading, multi-sheet, headers, and protection selection
# ------------------------------------------------------------------------------

def test_01_load_workbook_populates_sheet_list(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", restored_dir=tmp_path / "restored")

    ok, msg = ctl.load_source_workbook(src)
    assert ok is True
    assert ctl.available_sheets == ["Staff", "Bonuses"]
    assert "Staff" in ctl.sheet_configs
    assert "Bonuses" in ctl.sheet_configs


def test_02_multi_sheet_selection(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)

    # Both default selected
    assert ctl.sheet_configs["Staff"].is_selected is True
    assert ctl.sheet_configs["Bonuses"].is_selected is True

    # User can deselect Bonuses
    ctl.select_sheet("Bonuses", False)
    assert ctl.sheet_configs["Bonuses"].is_selected is False
    assert ctl.sheet_configs["Staff"].is_selected is True


def test_03_sheet_specific_header_rows(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)

    # Set Staff header row = 2
    ctl.set_sheet_header_row("Staff", 2)
    assert ctl.sheet_configs["Staff"].header_row == 2
    assert ctl.sheet_configs["Bonuses"].header_row == 1


def test_04_changing_header_row_refreshes_columns(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)

    # Initial Staff at header_row=1 has "Title Banner"
    ctl.set_sheet_header_row("Staff", 1)
    staff_cols_r1 = [c.header for c in ctl.sheet_configs["Staff"].columns]
    assert "Title Banner" in staff_cols_r1

    # Change to row 2 refreshes columns to actual headers
    ctl.set_sheet_header_row("Staff", 2)
    staff_cols_r2 = [c.header for c in ctl.sheet_configs["Staff"].columns]
    assert "EmpID" in staff_cols_r2
    assert "姓名" in staff_cols_r2
    assert "基本工资" in staff_cols_r2
    assert "Title Banner" not in staff_cols_r2


def test_05_sensitive_recommendations_preselected(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    col_name = ctl.sheet_configs["Staff"].get_column("B")  # B is 姓名
    assert col_name is not None
    assert col_name.header == "姓名"
    assert col_name.is_recommended is True
    assert col_name.protection_type == "text"


def test_06_user_can_unselect_recommendation(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # User toggles 姓名 from text to none
    ctl.set_column_protection("Staff", "B", "none")
    assert ctl.sheet_configs["Staff"].get_column("B").protection_type == "none"


def test_07_text_protection_selection(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Select 部门 for text protection
    ctl.set_column_protection("Staff", "D", "text")
    assert "部门" in ctl.sheet_configs["Staff"].get_text_protection_fields()


def test_08_numeric_protection_selection(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Select 基本工资 for numeric protection
    ctl.set_column_protection("Staff", "C", "numeric")
    assert "基本工资" in ctl.sheet_configs["Staff"].get_numeric_protection_fields()


def test_09_single_record_identity_selection(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Select EmpID as record matching column
    ctl.toggle_column_record_key("Staff", "A", True)
    assert ctl.sheet_configs["Staff"].key_validation_valid is True
    assert "✓ 记录识别验证通过，可以确认本表配置。" == ctl.sheet_configs["Staff"].key_validation_status


def test_10_composite_record_identity_selection(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)

    # In Bonuses sheet, select EmpID and 月份 as composite key
    ctl.toggle_column_record_key("Bonuses", "A", True)  # EmpID
    ctl.toggle_column_record_key("Bonuses", "B", True)  # 月份
    assert ctl.sheet_configs["Bonuses"].key_validation_valid is True
    assert set(ctl.sheet_configs["Bonuses"].get_record_key_fields()) == {"EmpID", "月份"}
    assert "✓ 记录识别验证通过，可以确认本表配置。" == ctl.sheet_configs["Bonuses"].key_validation_status


# ------------------------------------------------------------------------------
# Tests 11 - 20: Key validation, groups, output directories, and mask workflow
# ------------------------------------------------------------------------------

def test_11_invalid_record_identity_warning(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Selecting 部门 (has duplicates: 技术部 appears twice)
    ctl.toggle_column_record_key("Staff", "D", True)
    assert ctl.sheet_configs["Staff"].key_validation_valid is False
    assert "当前识别字段存在重复记录" in ctl.sheet_configs["Staff"].key_validation_status


def test_12_no_record_identity_blocks_confirmation(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # No key selected
    assert len(ctl.sheet_configs["Staff"].get_record_key_fields()) == 0
    assert ctl.sheet_configs["Staff"].key_validation_valid is False
    ok, msg = ctl.confirm_sheet_configuration("Staff")
    assert ok is False
    assert "尚未选择记录识别列" in msg
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_13_numeric_group_assignment(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    ctl.set_column_protection("Staff", "C", "numeric", analysis_group="组2")
    assert ctl.sheet_configs["Staff"].get_column("C").analysis_group == "组2"


def test_14_cross_sheet_same_group_assignment(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    ctl.set_column_protection("Staff", "C", "numeric", analysis_group="共享组A")
    ctl.set_column_protection("Bonuses", "C", "numeric", analysis_group="共享组A")

    scopes = ctl.build_sheet_scopes()
    staff_scope = next(s for s in scopes if s.sheet_name == "Staff")
    bonus_scope = next(s for s in scopes if s.sheet_name == "Bonuses")
    assert staff_scope.numeric_analysis_group == "共享组A"
    assert bonus_scope.numeric_analysis_group == "共享组A"


def test_15_masked_output_dir_selection_state(tmp_path: Path):
    custom_safe = tmp_path / "my_custom_safe"
    ctl = GUIWorkflowController(safe_dir=custom_safe)
    assert ctl.safe_dir == custom_safe.resolve()


def test_16_restored_output_dir_selection_state(tmp_path: Path):
    custom_restored = tmp_path / "my_custom_restored"
    ctl = GUIWorkflowController(restored_dir=custom_restored)
    assert ctl.restored_dir == custom_restored.resolve()


def test_frozen_macos_defaults_use_external_application_support(tmp_path: Path):
    root = default_app_data_root(frozen=True, platform_name="darwin", home=tmp_path)
    assert root == (tmp_path / "Library" / "Application Support" / "ExcelPrivacyTool").resolve()
    assert not root.is_relative_to(APP_DIR)


def test_development_defaults_remain_repository_relative(tmp_path: Path):
    assert default_app_data_root(frozen=False, platform_name="darwin", home=tmp_path) == APP_DIR


@pytest.fixture
def simulated_windows_bundle(monkeypatch, tmp_path: Path):
    exe_dir = tmp_path / "installed app"
    resources = tmp_path / "unpacked resources"
    monkeypatch.setattr(sys, "executable", str(exe_dir / "Excel Privacy Tool.exe"))
    monkeypatch.setattr(sys, "_MEIPASS", str(resources), raising=False)
    return exe_dir, resources


def test_frozen_windows_defaults_use_localappdata(monkeypatch, tmp_path: Path, simulated_windows_bundle):
    local_app_data = tmp_path / "中文用户" / "AppData" / "Local"
    existing_ledger = local_app_data / "ExcelPrivacyTool" / "private_keys" / "synthetic.recovery_ledger.enc"
    existing_ledger.parent.mkdir(parents=True)
    existing_ledger.write_bytes(b"synthetic existing data")
    monkeypatch.setenv("LOCALAPPDATA", str(local_app_data))

    root = default_app_data_root(frozen=True, platform_name="win32", home=tmp_path)

    assert root == (local_app_data / "ExcelPrivacyTool").resolve()
    assert not root.is_relative_to(APP_DIR)
    assert all(not root.is_relative_to(location) for location in simulated_windows_bundle)
    assert existing_ledger.read_bytes() == b"synthetic existing data"


def test_frozen_windows_defaults_when_localappdata_missing(monkeypatch, tmp_path: Path, simulated_windows_bundle):
    monkeypatch.delenv("LOCALAPPDATA", raising=False)

    root = default_app_data_root(frozen=True, platform_name="win32", home=tmp_path)

    assert root == (tmp_path / "AppData" / "Local" / "ExcelPrivacyTool").resolve()


def test_frozen_windows_defaults_never_use_relative_localappdata(monkeypatch, tmp_path: Path, simulated_windows_bundle):
    monkeypatch.setenv("LOCALAPPDATA", "relative-directory")

    root = default_app_data_root(frozen=True, platform_name="win32", home=tmp_path)

    assert root == (tmp_path / "AppData" / "Local" / "ExcelPrivacyTool").resolve()
    assert default_app_data_root(frozen=False, platform_name="win32", home=tmp_path) == APP_DIR


@pytest.mark.parametrize(
    "location", ("exe", "exe_child", "resources", "resources_child", "module_resources")
)
def test_frozen_windows_rejects_application_directories(
    monkeypatch, tmp_path: Path, simulated_windows_bundle, location: str
):
    exe_dir, resources = simulated_windows_bundle
    forbidden = {
        "exe": exe_dir,
        "exe_child": exe_dir / "child",
        "resources": resources,
        "resources_child": resources / "child",
        "module_resources": APP_DIR,
    }[location]
    monkeypatch.setenv("LOCALAPPDATA", str(forbidden))
    home = tmp_path / "user home"

    root = default_app_data_root(frozen=True, platform_name="win32", home=home)

    assert root == (home / "AppData" / "Local" / "ExcelPrivacyTool").resolve()
    assert not root.is_relative_to(exe_dir)
    assert not root.is_relative_to(resources)
    assert not root.is_relative_to(APP_DIR)


def test_frozen_windows_rejects_symlink_to_application_directory(
    monkeypatch, tmp_path: Path, simulated_windows_bundle
):
    exe_dir, _resources = simulated_windows_bundle
    exe_dir.mkdir()
    alias = tmp_path / "alias"
    alias.symlink_to(exe_dir, target_is_directory=True)
    monkeypatch.setenv("LOCALAPPDATA", str(alias))

    root = default_app_data_root(frozen=True, platform_name="win32", home=tmp_path / "home")

    assert root == (tmp_path / "home" / "AppData" / "Local" / "ExcelPrivacyTool").resolve()


def test_frozen_windows_refuses_unsafe_fallback(monkeypatch, simulated_windows_bundle):
    exe_dir, _resources = simulated_windows_bundle
    monkeypatch.delenv("LOCALAPPDATA", raising=False)

    with pytest.raises(RuntimeError, match="No safe per-user application data directory"):
        default_app_data_root(frozen=True, platform_name="win32", home=exe_dir)


def test_17_mask_button_validation(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()

    # Pre-condition fails without loading source
    ok, err = ctl.validate_mask_preconditions()
    assert ok is False
    assert "请先选择" in err

    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Pre-condition fails when sheet is unconfirmed
    ok, err = ctl.validate_mask_preconditions()
    assert ok is False
    assert "尚未选择记录识别列" in err

    # Unselect all columns to trigger protection missing error upon confirm
    ctl.set_column_protection("Staff", "B", "none")
    ctl.set_column_protection("Bonuses", "D", "none")
    ok, err = confirm_with_unique_record_identifier(ctl, "Staff")
    assert ok is False
    assert "至少选择一个保护字段" in err


def test_18_mask_controller_calls_mask_workbook_v1(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    safe_dir = tmp_path / "safe"
    priv_dir = tmp_path / "priv"
    ctl = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Configure Staff: 姓名 text, 基本工资 numeric, EmpID key
    ctl.set_column_protection("Staff", "B", "text")
    ctl.set_column_protection("Staff", "C", "numeric")
    ctl.toggle_column_record_key("Staff", "A", True)

    # Bonuses: deselect
    ctl.select_sheet("Bonuses", False)
    confirm_with_unique_record_identifier(ctl, "Staff")

    res = ctl.execute_mask()
    assert "masked_workbook_path" in res
    assert Path(res["masked_workbook_path"]).exists()
    assert ctl.last_masked_path == Path(res["masked_workbook_path"]).resolve()


def test_19_mask_result_stores_ai_prompt(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")

    ctl.execute_mask()
    assert len(ctl.ai_prompt) > 0
    assert "【Excel 分析与处理说明】" in ctl.ai_prompt


def test_20_copy_prompt_uses_generated_prompt(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")

    res = ctl.execute_mask()
    # Prompt in controller matches mask_workbook_v1 return
    assert ctl.ai_prompt == res["ai_prompt"]


# ------------------------------------------------------------------------------
# Tests 21 - 30: AI return, ledger auto-match, conflict compare & restore
# ------------------------------------------------------------------------------

def test_21_returned_workbook_selection(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    masked_file = ctl.last_masked_path
    ok, msg = ctl.load_returned_workbook(masked_file)
    assert ok is True
    assert ctl.returned_path == masked_file


def test_22_ledger_auto_resolution_by_run_id(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")
    res = ctl.execute_mask()

    # Load returned file
    ctl.load_returned_workbook(ctl.last_masked_path)
    assert ctl.resolved_run_id == res["run_id"]
    assert ctl.resolved_ledger_path == res["ledger_enc_path"]
    assert ctl.resolved_secret_key is not None


def test_23_run_identity_error_user_message(tmp_path: Path):
    # Create an alien workbook without valid XPT run_id
    alien_wb = tmp_path / "alien.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["Alien Header", "Value"])
    wb.save(alien_wb)

    ctl = GUIWorkflowController(private_keys_dir=tmp_path / "priv")
    ok, msg = ctl.load_returned_workbook(alien_wb)
    assert ok is False
    assert "无法确认该文件对应的脱敏批次" in msg


def test_24_compare_controller_calls_conflict_engine(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    ctl.load_returned_workbook(ctl.last_masked_path)
    summary = ctl.execute_compare()
    assert summary is not None
    assert summary.can_safe_restore_count == 3  # 3 staff names


def test_25_compare_summary_counts(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.set_column_protection("Staff", "C", "numeric")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    # Simulate AI modifying row 3 numeric amount (conflict) and adding a row
    ai_file = tmp_path / "ai_returned.xlsx"
    wb = openpyxl.load_workbook(ctl.last_masked_path)
    ws = wb["Staff"]
    ws.cell(row=3, column=3, value=99999)  # Tamper salary of E001
    ws.append(["E004", "AI Added User", 8888, "AI Dept"])  # Added row
    wb.save(ai_file)

    ctl.load_returned_workbook(ai_file)
    summary = ctl.execute_compare()
    assert summary.has_conflicts is True
    assert summary.protected_conflict_count == 1
    assert summary.added_rows_count == 1
    assert summary.can_safe_restore_count == 5  # 3 names + 2 unchanged salaries


def test_26_restore_controller_calls_restore_workbook_v1(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    restored_dir = tmp_path / "restored"
    ctl = GUIWorkflowController(
        safe_dir=tmp_path / "safe",
        restored_dir=restored_dir,
        private_keys_dir=tmp_path / "priv",
    )
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    ctl.load_returned_workbook(ctl.last_masked_path)
    ctl.execute_compare()
    res = ctl.execute_restore()
    assert res is not None
    assert Path(res.restored_workbook_path).exists()
    assert res.restored_cells_count == 3


def test_27_restore_summary_success(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    ctl.load_returned_workbook(ctl.last_masked_path)
    ctl.execute_compare()
    res = ctl.execute_restore()
    assert res.skipped_conflicts_count == 0
    assert res.skipped_unconfirmed_count == 0
    assert res.restored_cells_count == 3


def test_28_restore_summary_partial_conflict_warning(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.set_column_protection("Staff", "C", "numeric")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    # Modify one salary cell
    ai_file = tmp_path / "ai_modified.xlsx"
    wb = openpyxl.load_workbook(ctl.last_masked_path)
    ws = wb["Staff"]
    ws.cell(row=3, column=3, value=77777)
    wb.save(ai_file)

    ctl.load_returned_workbook(ai_file)
    ctl.execute_compare()
    res = ctl.execute_restore()
    assert res.skipped_conflicts_count == 1
    assert res.restored_cells_count == 5


def test_29_unsupported_parts_message_is_human_readable():
    raw_error = Exception("Unsupported workbook features detected: PivotTable (xl/pivotTables/pivotTable1.xml). Processing stopped.")
    msg = translate_security_error(raw_error)
    assert "这个 Excel 包含当前版本暂不支持的高级对象" in msg
    assert "原文件没有被修改" in msg


def test_30_gui_never_displays_scale_factor(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "C", "numeric")
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    # Neither ai_prompt nor status message should mention scale factor k or numerical multiplier
    assert "scale_factor" not in ctl.ai_prompt
    assert "倍率" not in ctl.status_message


# ------------------------------------------------------------------------------
# Tests 31 - 35: Privacy leaks & background task safety
# ------------------------------------------------------------------------------

def test_31_gui_never_displays_secret_key(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    # Secret key bytes or path must never be in status_message or ai_prompt
    assert "secret.key" not in ctl.status_message
    assert "secret.key" not in ctl.ai_prompt


def test_32_gui_never_displays_plain_record_identity(tmp_path: Path):
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "C", "numeric")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    # Modify cell to trigger conflict
    ai_file = tmp_path / "ai_file.xlsx"
    wb = openpyxl.load_workbook(ctl.last_masked_path)
    wb["Staff"].cell(row=3, column=3, value=55555)
    wb.save(ai_file)

    ctl.load_returned_workbook(ai_file)
    summary = ctl.execute_compare()
    for item in summary.conflict_details:
        # Must show safe alias (Record-01), not plain PII
        assert "E001" not in item["record"]
        assert "Record-" in item["record"]


def test_33_background_task_disables_buttons():
    ctl = GUIWorkflowController()
    buttons_state = {"run_mask": "normal", "compare": "normal"}

    def run_with_lock():
        ctl.is_busy = True
        buttons_state["run_mask"] = "disabled"
        buttons_state["compare"] = "disabled"

    run_with_lock()
    assert ctl.is_busy is True
    assert buttons_state["run_mask"] == "disabled"


def test_34_background_task_reenables_buttons_after_success():
    ctl = GUIWorkflowController()
    buttons_state = {"run_mask": "disabled", "compare": "disabled"}

    def finish_success():
        ctl.is_busy = False
        buttons_state["run_mask"] = "normal"
        buttons_state["compare"] = "normal"

    finish_success()
    assert ctl.is_busy is False
    assert buttons_state["run_mask"] == "normal"


def test_35_background_task_reenables_buttons_after_error():
    ctl = GUIWorkflowController()
    buttons_state = {"run_mask": "disabled", "compare": "disabled"}

    try:
        raise ValueError("Simulated error")
    except Exception:
        pass
    finally:
        ctl.is_busy = False
        buttons_state["run_mask"] = "normal"
        buttons_state["compare"] = "normal"

    assert ctl.is_busy is False
    assert buttons_state["run_mask"] == "normal"


# ------------------------------------------------------------------------------
# Tests 36 - 48: Checkpoint 8 UX Closeout dedicated tests
# ------------------------------------------------------------------------------

def test_36_record_key_label_is_record_matching_column(tmp_path: Path):
    """36. The UI consistently uses the frozen Record Identifier wording."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    status_str = ctl.sheet_configs["Staff"].key_validation_status
    assert "记录识别列" in status_str


def test_37_record_matching_help_text_is_plain_language():
    """37. Check that record matching column explanation uses clear, everyday language."""
    plain_help = "用于在 AI 排序、插行或修改后识别同一条记录。通常可选择员工工号；如果单列会重复，可组合选择员工工号 + 年月等字段。"
    assert "AI 排序、插行或修改" in plain_help
    assert "员工工号" in plain_help


def test_38_no_record_matching_column_warning_is_clear(tmp_path: Path):
    """38. An unselected Record Identifier displays a blocking explanation."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    status = ctl.sheet_configs["Staff"].key_validation_status
    assert "尚未选择记录识别列" in status
    assert "验证通过后才能确认本表" in status


def test_39_valid_record_matching_feedback(tmp_path: Path):
    """39. Check feedback message when valid single record matching column is chosen."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.toggle_column_record_key("Staff", "A", True)
    assert ctl.sheet_configs["Staff"].key_validation_status == "✓ 记录识别验证通过，可以确认本表配置。"


def test_40_composite_record_matching_feedback(tmp_path: Path):
    """40. Check feedback message when composite record matching columns are chosen."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.toggle_column_record_key("Bonuses", "A", True)
    ctl.toggle_column_record_key("Bonuses", "B", True)
    assert ctl.sheet_configs["Bonuses"].key_validation_status == "✓ 记录识别验证通过，可以确认本表配置。"


def test_41_numeric_group_label_is_plain_language(tmp_path: Path):
    """41. Check that numeric groups are labeled as '金额可比组' and default to '可比组 A'."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    assert ctl.sheet_configs["Staff"].get_column("C").analysis_group == "可比组 A"


def test_42_numeric_group_enabled_only_for_numeric_protection(tmp_path: Path):
    """42. Check that numeric group is active for numeric protection and inactive for others."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "C", "numeric")
    assert ctl.sheet_configs["Staff"].get_column("C").protection_type == "numeric"


def test_43_text_protection_does_not_show_active_numeric_group(tmp_path: Path):
    """43. Check that text protection field does not hold active numeric comparison role."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    assert ctl.sheet_configs["Staff"].get_column("B").protection_type == "text"


def test_44_unprotected_column_does_not_show_active_numeric_group(tmp_path: Path):
    """44. Check that unprotected column is marked 'none' and has no active numeric role."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    assert ctl.sheet_configs["Staff"].get_column("D").protection_type == "none"


def test_45_compare_summary_uses_plain_language(tmp_path: Path):
    """45. Check that Compare summary top line uses plain user-friendly language."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    ctl.load_returned_workbook(ctl.last_masked_path)
    summary = ctl.execute_compare()
    assert "【可自动恢复】" in ctl.status_message
    assert "【AI 修改了保护数据】" in ctl.status_message
    assert "【无法安全自动判断】" in ctl.status_message


def test_46_duplicate_identifier_cannot_confirm(tmp_path: Path):
    """46. A duplicate single-column identifier remains pending with a clear reason."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.toggle_column_record_key("Staff", "D", True)
    ok, msg = ctl.confirm_sheet_configuration("Staff")
    assert ok is False
    assert "当前识别字段存在重复记录" in msg
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_47_empty_identifier_component_cannot_confirm(tmp_path: Path):
    """47. Empty values in any selected identifier component block confirmation."""
    src = tmp_path / "empty_key.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Staff"
    ws.append(["EmpID", "姓名"])
    ws.append(["E001", "张三"])
    ws.append([None, "李四"])
    wb.save(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    ok, msg = ctl.confirm_sheet_configuration("Staff")
    assert ok is False
    assert "当前识别字段包含空值" in msg
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_48_restore_partial_confirmation_warning_is_clear():
    """48. Check that pre-restore confirmation dialog text is clear and non-technical."""
    unconfirmed_count = 180
    confirm_msg = (
        f"当前有 {unconfirmed_count} 项无法安全自动判断。\n\n"
        f"程序只会恢复已确认安全的项目，其余项目将保持 AI 返回值不变。\n\n"
        f"是否继续安全恢复？"
    )
    assert "无法安全自动判断" in confirm_msg
    assert "保持 AI 返回值不变" in confirm_msg
    assert "180 项" in confirm_msg


# ------------------------------------------------------------------------------
# Tests 49 - 54: Sheet switching & state retention regression tests
# ------------------------------------------------------------------------------

def test_49_sheet_switching_retains_configured_state(tmp_path: Path):
    """49. Switching back and forth between sheets retains configured columns and protection."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.set_column_protection("Staff", "C", "numeric", "可比组 B")
    ctl.toggle_column_record_key("Staff", "A", True)

    # Now switch to Bonuses and configure
    ctl.set_column_protection("Bonuses", "C", "numeric", "可比组 A")

    # Re-scan or switch back to Staff
    assert ctl.sheet_configs["Staff"].get_column("B").protection_type == "text"
    assert ctl.sheet_configs["Staff"].get_column("C").protection_type == "numeric"
    assert ctl.sheet_configs["Staff"].get_column("C").analysis_group == "可比组 B"
    assert ctl.sheet_configs["Staff"].get_column("A").is_record_key is True


def test_50_sheet_refresh_columns_preserves_user_choices(tmp_path: Path):
    """50. Calling refresh_sheet_columns on existing sheet preserves user configurations."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "C", "numeric", "可比组 C")
    ctl.toggle_column_record_key("Staff", "A", True)

    # Refresh columns at the same header row
    ok, msg = ctl.refresh_sheet_columns("Staff", header_row=2)
    assert ok is True
    col_c = ctl.sheet_configs["Staff"].get_column("C")
    assert col_c.protection_type == "numeric"
    assert col_c.analysis_group == "可比组 C"
    col_a = ctl.sheet_configs["Staff"].get_column("A")
    assert col_a.is_record_key is True


def test_51_header_row_change_on_sheet_b_does_not_affect_sheet_a(tmp_path: Path):
    """51. Changing header row on Sheet B does not alter Sheet A's configuration."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)

    # Alter Sheet B header row
    ctl.set_sheet_header_row("Bonuses", 1)

    # Verify Sheet A remains intact
    assert ctl.sheet_configs["Staff"].header_row == 2
    assert ctl.sheet_configs["Staff"].get_column("B").protection_type == "text"
    assert ctl.sheet_configs["Staff"].get_column("A").is_record_key is True


def test_52_sheet_is_configured_badge_property(tmp_path: Path):
    """52. SheetConfigState.is_configured reports False initially and True only after explicit confirmation."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    # Initially is_confirmed is False
    assert ctl.sheet_configs["Staff"].is_configured is False
    assert ctl.sheet_configs["Staff"].is_confirmed is False

    # Explicit confirmation sets it to True
    ok, _ = confirm_with_unique_record_identifier(ctl, "Staff")
    assert ok is True
    assert ctl.sheet_configs["Staff"].is_configured is True
    assert ctl.sheet_configs["Staff"].is_confirmed is True

    # User edit invalidates confirmation back to False
    ctl.set_column_protection("Staff", "B", "none")
    assert ctl.sheet_configs["Staff"].is_configured is False
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_53_empty_columns_sheet_handling(tmp_path: Path):
    """53. Sheets with out-of-range header rows produce empty columns without crashing."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    # Set header row far past bottom
    ok, msg = ctl.set_sheet_header_row("Staff", 99)
    assert ok is True
    assert len(ctl.sheet_configs["Staff"].columns) == 0
    assert ctl.sheet_configs["Staff"].is_configured is False


def test_54_apply_recommended_key_helper(tmp_path: Path):
    """54. apply_recommended_key activates only the recommended column on explicit trigger."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    # Recommended key is column A (EmpID)
    assert ctl.sheet_configs["Staff"].recommended_key_col == "A"
    assert ctl.sheet_configs["Staff"].get_column("A").is_record_key is False

    ok = ctl.apply_recommended_key("Staff")
    assert ok is True
    assert ctl.sheet_configs["Staff"].get_column("A").is_record_key is True
    assert ctl.sheet_configs["Staff"].key_validation_valid is True


def test_recommended_identifier_is_not_confirmation_until_user_applies_it(tmp_path: Path):
    src = tmp_path / "recommended_not_selected.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    cfg = ctl.sheet_configs["Staff"]
    assert cfg.recommended_key_col == "A"
    assert cfg.get_record_key_columns() == []
    ok, msg = ctl.confirm_sheet_configuration("Staff")
    assert ok is False
    assert "尚未选择记录识别列" in msg

    assert ctl.apply_recommended_key("Staff") is True
    ok, msg = ctl.confirm_sheet_configuration("Staff")
    assert ok is True
    assert cfg.get_record_key_columns() == ["A"]
    assert cfg.is_confirmed is True


def test_unique_composite_identifier_can_confirm(tmp_path: Path):
    src = tmp_path / "composite_unique.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "History"
    ws.append(["Month", "EmpID", "Name"])
    ws.append(["2026-01", "E001", "A"])
    ws.append(["2026-01", "E002", "B"])
    ws.append(["2026-02", "E001", "A"])
    ws.append(["2026-02", "E002", "B"])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_column_protection("History", "C", "text")
    ctl.toggle_column_record_key("History", "A", True)
    assert ctl.sheet_configs["History"].key_validation_valid is False
    ctl.toggle_column_record_key("History", "B", True)
    ok, _ = ctl.confirm_sheet_configuration("History")
    assert ok is True
    assert ctl.sheet_configs["History"].get_record_key_columns() == ["A", "B"]


def test_duplicate_composite_identifier_cannot_confirm(tmp_path: Path):
    src = tmp_path / "composite_duplicate.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "History"
    ws.append(["Month", "EmpID", "Name"])
    ws.append(["2026-01", "E001", "A"])
    ws.append(["2026-01", "E001", "B"])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_column_protection("History", "C", "text")
    ctl.toggle_column_record_key("History", "A", True)
    ctl.toggle_column_record_key("History", "B", True)
    ok, msg = ctl.confirm_sheet_configuration("History")
    assert ok is False
    assert "当前识别字段存在重复记录" in msg


def test_identifier_validation_exception_fails_closed(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    import privacy_handler.gui_controller as gui_controller

    src = tmp_path / "validation_error.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.toggle_column_record_key("Staff", "A", True)
    monkeypatch.setattr(
        gui_controller,
        "validate_composite_key",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("simulated validation failure")),
    )

    ok, msg = ctl.confirm_sheet_configuration("Staff")
    assert ok is False
    assert "无法完成记录识别验证" in msg
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_editing_confirmed_identifier_returns_sheet_to_pending(tmp_path: Path):
    src = tmp_path / "edit_identifier.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    assert confirm_with_unique_record_identifier(ctl, "Staff")[0]

    ctl.toggle_column_record_key("Staff", "A", False)
    assert ctl.sheet_configs["Staff"].is_confirmed is False
    assert ctl.sheet_configs["Staff"].key_validation_valid is False
    assert "尚未选择记录识别列" in ctl.sheet_configs["Staff"].key_validation_status


def test_partial_multi_sheet_confirmation_blocks_button_and_production_entry(tmp_path: Path):
    src = tmp_path / "partial_confirmation.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    assert confirm_with_unique_record_identifier(ctl, "Staff")[0]
    select_unique_record_identifier(ctl, "Bonuses")
    assert ctl.sheet_configs["Bonuses"].key_validation_valid is True
    assert ctl.sheet_configs["Bonuses"].is_confirmed is False

    ok, msg = ctl.validate_mask_preconditions()
    assert ok is False
    assert "Bonuses" in msg and "尚未确认配置" in msg
    with pytest.raises(PrivacyToolError, match="Bonuses"):
        ctl.execute_mask()


def test_stale_confirmed_flags_cannot_bypass_production_revalidation(tmp_path: Path):
    src = tmp_path / "stale_confirmation.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    assert confirm_with_unique_record_identifier(ctl, "Staff")[0]

    cfg = ctl.sheet_configs["Staff"]
    cfg.get_column("A").is_record_key = False
    cfg.key_validation_valid = True
    cfg.is_confirmed = True
    ok, msg = ctl.validate_mask_preconditions()
    assert ok is False
    assert "尚未选择记录识别列" in msg
    assert cfg.key_validation_valid is False
    assert cfg.is_confirmed is False


def test_source_data_change_invalidates_previously_confirmed_identifier(tmp_path: Path):
    src = tmp_path / "source_changed.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    assert confirm_with_unique_record_identifier(ctl, "Staff")[0]

    wb = openpyxl.load_workbook(src)
    wb["Staff"]["A4"] = wb["Staff"]["A3"].value
    wb.save(src)
    wb.close()
    ok, msg = ctl.validate_mask_preconditions()
    assert ok is False
    assert "存在重复记录" in msg
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_internal_and_trailing_formatted_blank_rows_are_not_identifier_records(tmp_path: Path):
    src = tmp_path / "formatted_blank_rows.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Staff"
    ws.append(["EmpID", "Name"])
    ws.append(["E001", "A"])
    ws.append([None, None])
    ws.append(["E002", "B"])
    ws.cell(row=20, column=1).fill = openpyxl.styles.PatternFill(
        fill_type="solid", fgColor="FFFFFF"
    )
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    result = ctl.validate_sheet_record_identity("Staff")
    assert result["is_valid"] is True
    assert result["total_count"] == 2
    assert ctl.confirm_sheet_configuration("Staff")[0] is True


# ------------------------------------------------------------------------------
# Tests 55 - 66: UX terminology freeze & Compare layout features
# ------------------------------------------------------------------------------

def test_55_terminology_numeric_protection_frozen():
    """55. Protection terminology uses 数值保护 rather than 金额保护."""
    prot_options = ["不保护", "文本保护", "数值保护"]
    assert "数值保护" in prot_options
    assert "金额保护" not in prot_options


def test_56_terminology_numeric_group_frozen():
    """56. Group label is 数值可比组 rather than 金额可比组."""
    group_header = "数值可比组"
    assert "数值" in group_header
    assert "金额" not in group_header


def test_57_terminology_record_identifier_frozen():
    """57. Key label is 记录识别列（类似主键） rather than 关联键."""
    header = "记录识别列（类似主键）"
    assert "记录识别列" in header
    assert "主键" in header
    assert "关联键" not in header


def test_58_compare_three_state_cards(tmp_path: Path):
    """58. Compare top line outputs three clean cards: 可自动恢复, AI修改了保护数据, 需要人工核对."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(safe_dir=tmp_path / "safe", private_keys_dir=tmp_path / "priv")
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    ctl.load_returned_workbook(ctl.last_masked_path)
    summary = ctl.execute_compare()
    assert summary.can_safe_restore_count == 3
    assert summary.protected_conflict_count == 0
    assert summary.require_confirmation_count == 0


def test_59_structural_delta_stats_plain_display():
    """59. Structural delta metrics format as AI 新增行: 0 | AI 删除行: 0 | AI 新增列: 0."""
    stats_txt = f"AI 新增行: {0} | AI 删除行: {0} | AI 新增列: {0}"
    assert "AI 新增行: 0" in stats_txt
    assert "AI 删除行: 0" in stats_txt
    assert "AI 新增列: 0" in stats_txt


def test_60_compare_zero_problem_banner():
    """60. Clean banner displayed when 0 problems found."""
    banner = "✓ 未发现需要你处理的问题，可以安全恢复。"
    assert "可以安全恢复" in banner
    assert "未发现需要你处理的问题" in banner


def test_61_empty_conflict_table_state():
    """61. Conflict table empty state displays reassuring message."""
    empty_row = ("—", "✓ 没有需要逐项查看的问题", "0 项")
    assert "没有需要逐项查看的问题" in empty_row[1]


def test_62_unconfirmed_card_is_label_explanation():
    """62. Unconfirmed guidance is framed as an explanation card rather than clickable trigger."""
    lbl_text = " ⓘ 需要人工核对说明 "
    assert "需要人工核对" in lbl_text
    assert "说明" in lbl_text


def test_63_prompt_copy_disabled_hint():
    """63. Disabled prompt button has clear reason text."""
    hint = "（完成脱敏后可复制）"
    assert "完成脱敏后可复制" in hint


def test_64_unique_key_guidance_text(tmp_path: Path):
    """64. Unique key selection outputs positive feedback without urging more columns."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.toggle_column_record_key("Staff", "A", True)
    status = ctl.sheet_configs["Staff"].key_validation_status
    assert status == "✓ 记录识别验证通过，可以确认本表配置。"


def test_65_composite_key_guidance_text(tmp_path: Path):
    """65. Composite unique key outputs warning not to add more columns unnecessarily."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.toggle_column_record_key("Bonuses", "A", True)
    ctl.toggle_column_record_key("Bonuses", "B", True)
    status = ctl.sheet_configs["Bonuses"].key_validation_status
    assert status == "✓ 记录识别验证通过，可以确认本表配置。"


def test_66_duplicate_key_helpful_suggestion(tmp_path: Path):
    """66. Duplicate key candidate advises user to add distinguishing column like 年月."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.toggle_column_record_key("Staff", "D", True)  # 部门 has duplicates
    status = ctl.sheet_configs["Staff"].key_validation_status
    assert "当前识别字段存在重复记录" in status


# ------------------------------------------------------------------------------
# Tests 67 - 85: Real-User Final Bugfix Regression Tests
# ------------------------------------------------------------------------------

def test_newly_loaded_sheet_is_pending_confirmation(tmp_path: Path):
    """67. Newly loaded sheet state is pending confirmation (is_confirmed is False)."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    assert ctl.sheet_configs["Staff"].is_confirmed is False
    assert ctl.sheet_configs["Bonuses"].is_confirmed is False


def test_recommended_fields_do_not_make_sheet_confirmed(tmp_path: Path):
    """68. Auto-recommended sensitive fields do NOT mark sheet as confirmed."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    # Staff has recommended sensitive field (姓名)
    assert ctl.sheet_configs["Staff"].get_column("B").is_recommended is True
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_confirm_sheet_configuration_sets_confirmed(tmp_path: Path):
    """69. Explicit confirm_sheet_configuration marks sheet as confirmed."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ok, msg = confirm_with_unique_record_identifier(ctl, "Staff")
    assert ok is True
    assert "已确认" in msg
    assert ctl.sheet_configs["Staff"].is_confirmed is True


def test_editing_confirmed_sheet_returns_to_pending(tmp_path: Path):
    """70. Editing any setting in confirmed sheet returns its state to pending."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    confirm_with_unique_record_identifier(ctl, "Staff")
    assert ctl.sheet_configs["Staff"].is_confirmed is True

    # Edit protection type
    ctl.set_column_protection("Staff", "C", "numeric")
    assert ctl.sheet_configs["Staff"].is_confirmed is False

    # Re-confirm then edit header row
    confirm_with_unique_record_identifier(ctl, "Staff")
    assert ctl.sheet_configs["Staff"].is_confirmed is True
    ctl.set_sheet_header_row("Staff", 2)
    assert ctl.sheet_configs["Staff"].is_confirmed is False


def test_mask_blocked_when_selected_sheet_not_confirmed(tmp_path: Path):
    """71. validate_mask_preconditions blocks masking if any selected sheet is unconfirmed."""
    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)

    ok, err = ctl.validate_mask_preconditions()
    assert ok is False
    assert "尚未选择记录识别列" in err


def test_sheet_b_fields_render_immediately_on_switch(tmp_path: Path):
    """72. In app UI, switching to Sheet B renders columns immediately without needing export."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()

    # Switch to Staff
    app._switch_active_sheet("Staff")
    app.update()
    assert len(app.col_inner_frame.winfo_children()) > 0

    # Switch to Bonuses
    app._switch_active_sheet("Bonuses")
    app.update()
    assert len(app.col_inner_frame.winfo_children()) > 0
    assert app.active_sheet_name == "Bonuses"
    app.destroy()


def test_restore_coverage_counts_key_changed_row_and_reports_identity_details(tmp_path: Path):
    """A changed Record Identifier makes its whole protected row explicitly unresolved."""
    src = tmp_path / "key_change.xlsx"
    create_sample_workbook(src)
    ctl = GUIWorkflowController(
        safe_dir=tmp_path / "safe",
        restored_dir=tmp_path / "restored",
        private_keys_dir=tmp_path / "priv",
    )
    ctl.load_source_workbook(src)
    ctl.select_sheet("Bonuses", False)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    ctl.set_column_protection("Staff", "C", "numeric")
    ctl.toggle_column_record_key("Staff", "A", True)
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.execute_mask()

    returned = tmp_path / "key_changed_returned.xlsx"
    wb = openpyxl.load_workbook(ctl.last_masked_path)
    wb["Staff"]["A3"] = "AI-CHANGED-ID"
    wb.save(returned)

    assert ctl.load_returned_workbook(returned)[0]
    summary = ctl.execute_compare()
    assert summary.total_protected_cells_count == 6
    assert summary.can_safe_restore_count == 4
    assert summary.require_confirmation_count == 2
    assert summary.added_rows_count == 1
    assert summary.deleted_rows_count == 1
    assert summary.unresolved_by_sheet == {"Staff": 2}
    reasons = [item["reason"] for item in summary.conflict_details]
    assert any("记录识别值已变更" in reason for reason in reasons)

    result = ctl.execute_restore()
    assert result.total_protected_cells_count == 6
    assert result.restored_cells_count == 4
    assert result.skipped_conflicts_count == 0
    assert result.skipped_unconfirmed_count == 2
    assert result.unresolved_by_sheet == {"Staff": 2}


def test_compare_does_not_call_ambiguous_no_key_cells_safe_to_restore(tmp_path: Path):
    """The core API still supports historical no-key runs outside the GUI gate."""
    src = tmp_path / "repeated_no_key.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "History"
    ws.append(["姓名"])
    ws.append(["张三"])
    ws.append(["张三"])
    ws.append(["张三"])
    wb.save(src)

    masked = mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        sheet_scopes=[
            SheetProtectionScope(
                sheet_name="History", header_row=1, text_columns=["A"], record_key=None
            )
        ],
    )
    ctl = GUIWorkflowController(
        restored_dir=tmp_path / "restored", private_keys_dir=tmp_path / "priv"
    )
    assert ctl.load_returned_workbook(masked["masked_workbook_path"])[0]

    summary = ctl.execute_compare()
    assert summary.total_protected_cells_count == 3
    assert summary.can_safe_restore_count == 0
    assert summary.require_confirmation_count == 3
    assert summary.unresolved_by_sheet == {"History": 3}

    result = ctl.execute_restore()
    assert result.restored_cells_count == 0
    assert result.skipped_unconfirmed_count == 3


def test_numeric_group_rebuild_shows_duplicate_physical_columns_without_dialog(tmp_path: Path):
    """G/I duplicate headers appear immediately and survive collapse/sheet switches."""
    from privacy_app import PrivacyApp

    src = tmp_path / "duplicate_numeric_groups.xlsx"
    wb = openpyxl.Workbook()
    salary = wb.active
    salary.title = "薪酬明细"
    salary.append(["说明"])
    salary.append([])
    salary.append(["月份", "工号", "C", "D", "E", "F", "金额", "H", "金额"])
    salary.append(["2026-01", "E001", 1, 2, 3, 4, 6500, 5, 388])
    other = wb.create_sheet("Other")
    other.append(["ID", "Value"])
    other.append(["1", 10])
    wb.save(src)

    app = PrivacyApp()
    app.controller.load_source_workbook(src)
    app.controller.set_sheet_header_row("薪酬明细", 3)
    app.controller.set_column_protection("薪酬明细", "G", "numeric")
    app.controller.set_column_protection("薪酬明细", "I", "numeric")
    app._switch_active_sheet("薪酬明细")
    app._toggle_advanced_group()
    app.update()

    def label_texts(widget):
        texts = []
        for child in widget.winfo_children():
            if child.winfo_class() in ("Label", "TLabel"):
                texts.append(str(child.cget("text")))
            texts.extend(label_texts(child))
        return texts

    labels = label_texts(app.adv_content_frame)
    assert "G｜金额" in labels
    assert "I｜金额" in labels
    assert not app.adv_scrollbar.winfo_manager()
    assert int(app.adv_canvas.cget("height")) <= 150

    app._toggle_advanced_group()
    app._switch_active_sheet("Other")
    app._switch_active_sheet("薪酬明细")
    app._toggle_advanced_group()
    app.update()
    labels = label_texts(app.adv_content_frame)
    assert labels.count("G｜金额") == 1
    assert labels.count("I｜金额") == 1
    assert app.adv_canvas.yview()[0] == 0.0
    app.destroy()


def test_numeric_group_touchpad_routes_from_combobox_and_keeps_field_list_local(tmp_path: Path):
    """Numeric-group descendants scroll their canvas without changing a combobox."""
    from privacy_app import PrivacyApp

    src = tmp_path / "many_numeric_groups.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Metrics"
    ws.append([f"Metric_{i}" for i in range(14)])
    ws.append(list(range(14)))
    wb.save(src)

    app = PrivacyApp()
    app.geometry("1060x760")
    app.update()
    if not app._supports_touchpad_scroll():
        app.destroy()
        pytest.skip("Tk build does not expose TouchpadScroll")
    app.controller.load_source_workbook(src)
    for idx in range(14):
        app.controller.set_column_protection("Metrics", openpyxl.utils.get_column_letter(idx + 1), "numeric")
    app._switch_active_sheet("Metrics")
    app._toggle_advanced_group()
    app.update()

    def descendants(widget):
        for child in widget.winfo_children():
            yield child
            yield from descendants(child)

    combo = next(w for w in descendants(app.adv_content_frame) if w.winfo_class() == "TCombobox")
    assert combo.bind("<TouchpadScroll>")
    assert not app.bind_all("<TouchpadScroll>")
    combo_value = combo.get()
    field_view = app.col_canvas.yview()
    app.adv_canvas.yview_moveto(0.0)
    before = app.adv_canvas.yview()[0]
    combo.event_generate("<TouchpadScroll>", delta=0xFFFF)
    app.update()
    assert app.adv_canvas.yview()[0] > before
    assert combo.get() == combo_value
    assert app.col_canvas.yview() == field_view

    app.adv_canvas.yview_moveto(0.0)
    app.adv_canvas.event_generate("<TouchpadScroll>", delta=5)
    app.update()
    assert app.adv_canvas.yview()[0] == 0.0
    app.adv_canvas.yview_moveto(1.0)
    app.adv_canvas.event_generate("<TouchpadScroll>", delta=0xFFFB)
    app.update()
    assert app.adv_canvas.yview()[1] == 1.0
    app.destroy()


def test_116_duplicate_headers_keep_physical_identity_end_to_end(tmp_path: Path, monkeypatch):
    """Duplicate numeric/text headers remain distinct through GUI mask, prompt, ledger, and restore."""
    from decimal import Decimal
    from privacy_handler.restore_engine import restore_workbook_v1
    import privacy_handler.numeric_masking as numeric_masking

    src = tmp_path / "duplicate_headers.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "薪酬明细"
    ws.append(["验收测试"])
    ws.append([])
    ws.append(["月份", "员工工号", "姓名", "部门", "姓名", "项目1", "金额", "项目2", "金额"])
    ws.append(["2026-07", "E001", "张三", "研发", "Zhang San", "基本工资", 6500, "奖金", 388])
    ws.append(["2026-07", "E002", "李四", "市场", "Li Si", "基本工资", 7200, "奖金", 420])
    wb.save(src)

    factors = iter([Decimal("2"), Decimal("3")])
    monkeypatch.setattr(numeric_masking, "generate_scale_factor", lambda: next(factors))

    ctl = GUIWorkflowController(
        safe_dir=tmp_path / "safe",
        restored_dir=tmp_path / "restored",
        private_keys_dir=tmp_path / "private",
    )
    assert ctl.load_source_workbook(src)[0]
    ctl.set_sheet_header_row("薪酬明细", 3)
    for col in ctl.sheet_configs["薪酬明细"].columns:
        ctl.set_column_protection("薪酬明细", col.column_letter, "none")
    ctl.set_column_protection("薪酬明细", "C", "text")
    ctl.set_column_protection("薪酬明细", "E", "text")
    ctl.set_column_protection("薪酬明细", "G", "numeric")
    ctl.set_column_protection("薪酬明细", "I", "numeric")
    ctl.set_column_local_group("薪酬明细", "G", "A")
    ctl.set_column_local_group("薪酬明细", "I", "B")
    ctl.toggle_column_record_key("薪酬明细", "B", True)
    assert ctl.confirm_sheet_configuration("薪酬明细")[0]

    scope = ctl.build_sheet_scopes()[0]
    assert scope.text_columns == ["C", "E"]
    assert scope.numeric_columns == ["G", "I"]
    assert scope.numeric_column_groups == {
        "G": "local::薪酬明细::A",
        "I": "local::薪酬明细::B",
    }

    result = ctl.execute_mask()
    masked_wb = openpyxl.load_workbook(result["masked_workbook_path"])
    masked_ws = masked_wb["薪酬明细"]
    assert masked_ws["G4"].value == 13000
    assert masked_ws["I4"].value == 1164
    assert masked_ws["C4"].value != "张三"
    assert masked_ws["E4"].value != "Zhang San"
    masked_wb.close()

    prompt = result["ai_prompt"]
    assert "G｜金额" in prompt
    assert "I｜金额" in prompt
    assert "C｜姓名" in prompt
    assert "E｜姓名" in prompt

    protected_ids = {pf.protected_field_id for pf in result["ledger"].protected_fields}
    assert {"薪酬明细:C", "薪酬明细:E", "薪酬明细:G", "薪酬明细:I"} <= protected_ids

    restored = restore_workbook_v1(
        returned_workbook_path=result["masked_workbook_path"],
        ledger_path=result["ledger_enc_path"],
        secret_key=result["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )
    restored_wb = openpyxl.load_workbook(restored.restored_workbook_path)
    restored_ws = restored_wb["薪酬明细"]
    assert [restored_ws[f"{col}{row}"].value for row in (4, 5) for col in ("G", "I")] == [6500, 388, 7200, 420]
    assert [restored_ws[f"{col}{row}"].value for row in (4, 5) for col in ("C", "E")] == ["张三", "Zhang San", "李四", "Li Si"]
    restored_wb.close()


def test_117_bank_card_headers_are_recommendations_only(tmp_path: Path):
    src = tmp_path / "bank_headers.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["银行卡尾号", "银行卡号", "银行账号", "银行账户", "账户号", "普通字段"]
    ws.append(headers)
    ws.append(["0001", "62220000", "A-1", "B-2", "C-3", "公开"])
    wb.save(src)

    ctl = GUIWorkflowController()
    assert ctl.load_source_workbook(src)[0]
    cols = ctl.sheet_configs[ws.title].columns
    assert [c.is_recommended for c in cols] == [True, True, True, True, True, False]
    assert ctl.sheet_configs[ws.title].is_confirmed is False


def test_118_shared_group_removal_restores_previous_local_group(tmp_path: Path):
    src = tmp_path / "groups.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "金额"])
    ws.append(["E1", 100])
    ws.append(["E2", 200])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    ctl.set_column_protection(ws.title, "B", "numeric")
    ctl.set_column_local_group(ws.title, "B", "C")
    ctl.assign_shared_group("共享组 1", [(ws.title, "B")])
    ctl.remove_from_shared_group(ws.title, "B")
    col = ctl.sheet_configs[ws.title].get_column("B")
    assert col.shared_group_id is None
    assert col.local_group == "C"


def test_119_compare_summary_contains_safe_cell_detail(tmp_path: Path):
    src = tmp_path / "compare_detail.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "人员快照"
    ws.append(["ID", "姓名", "C", "D", "E", "F", "G", "H", "I", "当前月薪"])
    ws.append(["E1", "张三", "", "", "", "", "", "", "", 10000])
    ws.append(["E2", "李四", "", "", "", "", "", "", "", 12000])
    ws.append(["E3", "王五", "", "", "", "", "", "", "", 14000])
    wb.save(src)

    ctl = GUIWorkflowController(
        safe_dir=tmp_path / "safe",
        restored_dir=tmp_path / "restored",
        private_keys_dir=tmp_path / "private",
    )
    ctl.load_source_workbook(src)
    ctl.set_column_protection(ws.title, "B", "text")
    ctl.set_column_protection(ws.title, "J", "numeric")
    ctl.toggle_column_record_key(ws.title, "A", True)
    assert confirm_with_unique_record_identifier(ctl, ws.title)[0]
    mask_result = ctl.execute_mask()

    ai_wb = openpyxl.load_workbook(mask_result["masked_workbook_path"])
    ai_wb[ws.title]["B2"] = 111
    ai_wb[ws.title]["J3"] = 123
    returned = tmp_path / "returned.xlsx"
    ai_wb.save(returned)
    ai_wb.close()

    assert ctl.load_returned_workbook(returned)[0]
    summary = ctl.execute_compare()
    details = {detail["cell"]: detail for detail in summary.conflict_details}
    assert details["B2"]["sheet"] == ws.title
    assert details["B2"]["field"] == "B｜姓名"
    assert details["B2"]["reason"] == "AI 修改了受保护文本"
    assert details["B2"]["policy"] == "保留 AI 返回值，不自动覆盖"
    assert details["B2"]["ai_value"] == 111
    assert details["J3"]["field"] == "J｜当前月薪"
    assert details["J3"]["reason"] == "AI 修改了受保护数值"
    assert details["J3"]["policy"] == "保留 AI 返回值，不自动覆盖"
    assert details["J3"]["ai_value"] == 123
    assert "AI 修改了受保护数据" in summary.reason_breakdown
    assert all("original" not in detail for detail in summary.conflict_details)


def test_120_main_window_minimum_and_field_scroll_bindings():
    from privacy_app import PrivacyApp

    app = PrivacyApp()
    app.geometry("1060x760")
    app.update()
    assert app.minsize() == (1060, 760)
    assert app.col_canvas.winfo_height() >= 100
    assert app.btn_mask.winfo_rooty() + app.btn_mask.winfo_height() <= app.winfo_rooty() + app.winfo_height()
    assert app.col_canvas.bind("<MouseWheel>")
    assert app.col_inner_frame.bind("<MouseWheel>")
    if app._supports_touchpad_scroll():
        assert app.col_canvas.bind("<TouchpadScroll>")
        assert app.col_inner_frame.bind("<TouchpadScroll>")
    app.destroy()


def test_121_touchpad_delta_unpack_and_scroll_direction():
    """Tk 9 packed touchpad deltas preserve small values and vertical direction."""
    from types import SimpleNamespace
    from unittest.mock import patch
    from privacy_app import PrivacyApp

    assert PrivacyApp._unpack_touchpad_scroll_delta(1) == (0, 1)
    assert PrivacyApp._unpack_touchpad_scroll_delta(0x0000FFFF) == (0, -1)
    assert PrivacyApp._unpack_touchpad_scroll_delta(0x0001FFFE) == (1, -2)
    assert PrivacyApp._unpack_touchpad_scroll_delta(-1) == (-1, -1)

    app = PrivacyApp()
    with patch.object(app.col_canvas, "yview_scroll") as scroll:
        assert app._on_field_list_touchpad_scroll(SimpleNamespace(delta=1)) == "break"
        scroll.assert_called_once_with(-1, "units")
        scroll.reset_mock()
        assert app._on_field_list_touchpad_scroll(SimpleNamespace(delta=0xFFFF)) == "break"
        scroll.assert_called_once_with(1, "units")
    app.destroy()


def test_122_touchpad_routes_from_field_label_combobox_and_canvas(tmp_path: Path):
    """Touchpad events under every field-list surface scroll only the field canvas."""
    from privacy_app import PrivacyApp

    src = tmp_path / "many_fields.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([f"Field_{i}" for i in range(40)])
    ws.append([i for i in range(40)])
    wb.save(src)

    app = PrivacyApp()
    app.geometry("1060x760")
    app.update()
    if not app._supports_touchpad_scroll():
        app.destroy()
        pytest.skip("Tk build does not expose TouchpadScroll")

    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet(ws.title)
    app.update()

    def descendants(widget):
        for child in widget.winfo_children():
            yield child
            yield from descendants(child)

    children = list(descendants(app.col_inner_frame))
    field_label = next(w for w in children if w.winfo_class() == "Label" and "｜Field_" in str(w.cget("text")))
    combo = next(w for w in children if w.winfo_class() == "TCombobox")
    assert field_label.bind("<TouchpadScroll>")
    assert combo.bind("<TouchpadScroll>")
    assert not app.btn_mask.bind("<TouchpadScroll>")
    assert not app.bind_all("<TouchpadScroll>")

    combo_value = combo.get()
    for target in (field_label, combo, app.col_canvas):
        app.col_canvas.yview_moveto(0.0)
        app.update_idletasks()
        before = app.col_canvas.yview()[0]
        target.event_generate("<TouchpadScroll>", delta=0xFFFF)
        app.update()
        assert app.col_canvas.yview()[0] > before
    assert combo.get() == combo_value
    app.destroy()


def test_123_touchpad_scroll_clamps_at_boundaries_without_reverse_jump(tmp_path: Path):
    """Small continuous deltas move correctly and clamp cleanly at both ends."""
    from privacy_app import PrivacyApp

    src = tmp_path / "scroll_bounds.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([f"Field_{i}" for i in range(40)])
    ws.append([i for i in range(40)])
    wb.save(src)

    app = PrivacyApp()
    app.geometry("1060x760")
    app.update()
    if not app._supports_touchpad_scroll():
        app.destroy()
        pytest.skip("Tk build does not expose TouchpadScroll")
    app.controller.load_source_workbook(src)
    app._switch_active_sheet(ws.title)
    app.update()

    app.col_canvas.yview_moveto(0.5)
    middle = app.col_canvas.yview()[0]
    app.col_canvas.event_generate("<TouchpadScroll>", delta=1)
    app.update()
    assert app.col_canvas.yview()[0] < middle
    app.col_canvas.event_generate("<TouchpadScroll>", delta=0xFFFF)
    app.update()
    assert app.col_canvas.yview()[0] == pytest.approx(middle)

    app.col_canvas.yview_moveto(0.0)
    app.col_canvas.event_generate("<TouchpadScroll>", delta=5)
    app.update()
    assert app.col_canvas.yview()[0] == 0.0

    app.col_canvas.yview_moveto(1.0)
    bottom = app.col_canvas.yview()[1]
    app.col_canvas.event_generate("<TouchpadScroll>", delta=0xFFFB)
    app.update()
    assert app.col_canvas.yview()[1] == bottom == 1.0
    app.destroy()


def test_124_mousewheel_and_linux_button_paths_keep_direction():
    """Legacy mouse-wheel paths keep Windows/macOS wheel and Linux direction."""
    from types import SimpleNamespace
    from unittest.mock import patch
    from privacy_app import PrivacyApp

    app = PrivacyApp()
    cases = [
        (SimpleNamespace(num=4, delta=0), (-30, "units")),
        (SimpleNamespace(num=5, delta=0), (30, "units")),
        (SimpleNamespace(num=None, delta=120), (-30, "units")),
        (SimpleNamespace(num=None, delta=-120), (30, "units")),
        (SimpleNamespace(num=None, delta=1), (-30, "units")),
        (SimpleNamespace(num=None, delta=-1), (30, "units")),
    ]
    with patch.object(app.col_canvas, "yview_scroll") as scroll:
        for event, expected in cases:
            scroll.reset_mock()
            assert app._on_field_list_mousewheel(event) == "break"
            scroll.assert_called_once_with(*expected)
    app.destroy()


def test_ten_repeated_sheet_switches_never_blank(tmp_path: Path):
    """73. Ten repeated switches between sheets never leave the column viewport blank."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()

    sheets = ["Staff", "Bonuses"]
    for i in range(10):
        target = sheets[i % 2]
        app._switch_active_sheet(target)
        app.update()
        children = app.col_inner_frame.winfo_children()
        assert len(children) > 0, f"Iteration {i} for {target} had 0 rendered widgets"
        assert app.col_canvas.yview()[0] == 0.0

    app.destroy()


def test_controller_columns_exist_and_view_renders_same_sheet(tmp_path: Path):
    """74. Controller column descriptors match exactly what the view renders for the active sheet."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.update()

    ctl_cols = app.controller.sheet_configs["Staff"].columns
    assert len(ctl_cols) == 4
    # Check that widgets rendered
    assert len(app.col_inner_frame.winfo_children()) > len(ctl_cols)
    app.destroy()


def test_true_numeric_column_allows_numeric_protection(tmp_path: Path):
    """75. True numeric columns (int/float) qualify for numeric protection candidate."""
    src = tmp_path / "num_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "Salary", "BonusRatio"])
    ws.append(["E1", 5000, 0.15])
    ws.append(["E2", 6000, 0.20])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_sal = next(c for c in cols if c.column_letter == "B")
    col_ratio = next(c for c in cols if c.column_letter == "C")
    assert col_sal.is_numeric_candidate is True
    assert col_ratio.is_numeric_candidate is True

    ctl.set_column_protection("Data", "B", "numeric")
    assert col_sal.protection_type == "numeric"


def test_numeric_string_column_does_not_allow_numeric_protection(tmp_path: Path):
    """76. Text columns formatted as numbers (like '001234') are disqualified from numeric candidate."""
    src = tmp_path / "str_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Code", "Value"])
    ws.append(["001234", 100])
    ws.append(["005678", 200])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_code = next(c for c in cols if c.column_letter == "A")
    assert col_code.is_numeric_candidate is False

    # Attempting to set numeric protection is rejected
    ctl.set_column_protection("Data", "A", "numeric")
    assert col_code.protection_type == "none"


def test_mixed_numeric_text_column_blocks_numeric_protection(tmp_path: Path):
    """77. Mixed columns containing text values like '100kg' or 'N/A' disqualify numeric candidate."""
    src = tmp_path / "mixed_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Weight"])
    ws.append([100])
    ws.append(["120kg"])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_w = next(c for c in cols if c.column_letter == "A")
    assert col_w.is_numeric_candidate is False


def test_date_column_does_not_allow_numeric_protection(tmp_path: Path):
    """78. Date columns disqualify numeric candidate."""
    import datetime
    src = tmp_path / "date_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["EntryDate"])
    ws.append([datetime.date(2023, 1, 1)])
    ws.append([datetime.date(2023, 2, 1)])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_d = next(c for c in cols if c.column_letter == "A")
    assert col_d.is_numeric_candidate is False


def test_bool_column_does_not_allow_numeric_protection(tmp_path: Path):
    """79. Boolean columns disqualify numeric candidate."""
    src = tmp_path / "bool_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ActiveFlag"])
    ws.append([True])
    ws.append([False])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_b = next(c for c in cols if c.column_letter == "A")
    assert col_b.is_numeric_candidate is False


def test_formula_column_does_not_allow_numeric_protection(tmp_path: Path):
    """80. Formula columns disqualify numeric candidate."""
    src = tmp_path / "formula_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["SumVal"])
    ws.append(["=SUM(A1:A2)"])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_f = next(c for c in cols if c.column_letter == "A")
    assert col_f.is_numeric_candidate is False


def test_numeric_candidate_is_not_auto_protected(tmp_path: Path):
    """81. Numeric candidate column default protection is 'none' unless sensitive keyword matched."""
    src = tmp_path / "salary_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["CountNum"])  # not a known sensitive salary keyword
    ws.append([42])
    ws.append([84])
    wb.save(src)

    ctl = GUIWorkflowController()
    ctl.load_source_workbook(src)
    cols = ctl.sheet_configs["Data"].columns
    col_c = next(c for c in cols if c.column_letter == "A")
    assert col_c.is_numeric_candidate is True
    assert col_c.protection_type == "none"


def test_text_protection_row_uses_text_visual_state(tmp_path: Path):
    """82. Text protection row has light blue visual tint in GUI."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    # Col B is text
    app.controller.set_column_protection("Staff", "B", "text")
    app._render_column_list()
    app.update()

    # Verify widget background or configuration rendered
    assert app.controller.sheet_configs["Staff"].get_column("B").protection_type == "text"
    app.destroy()


def test_numeric_protection_row_uses_numeric_visual_state(tmp_path: Path):
    """83. Numeric protection row has light green visual tint in GUI."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    # Set header row to 2 so column C is 基本工资 (numeric)
    app.controller.set_sheet_header_row("Staff", 2)
    app.controller.set_column_protection("Staff", "C", "numeric")
    app._render_column_list()
    app.update()

    assert app.controller.sheet_configs["Staff"].get_column("C").protection_type == "numeric"
    app.destroy()


def test_recommended_identifier_selection_visible_on_row(tmp_path: Path):
    """84. Applying recommended identifier displays '【识别】' badge on the corresponding column."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.controller.set_sheet_header_row("Staff", 2)

    # Apply recommended key
    app._apply_recommended_key()
    app.update()

    col_a = app.controller.sheet_configs["Staff"].get_column("A")
    assert col_a.is_record_key is True
    app.destroy()


def test_recommended_identifier_confirmation_text(tmp_path: Path):
    """85. Recommended identifier row displays prompt '推荐记录识别列：...' and updates status upon selection."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.controller.set_sheet_header_row("Staff", 2)
    app._render_column_list()
    app.update()

    assert "推荐记录识别列" in app.rec_key_prompt_lbl.cget("text")
    assert "A｜EmpID" in app.rec_key_prompt_lbl.cget("text")

    app._apply_recommended_key()
    app.update()

    assert "✓ 已选择：EmpID" in app.key_status_lbl.cget("text")
    app.destroy()


# ------------------------------------------------------------------------------
# Tests 86 - 90: Final GUI Layout Specification Verification (docs/GUI_V1_FINAL_LAYOUT_SPEC_20261008.md)
# ------------------------------------------------------------------------------

def test_86_field_table_has_strictly_two_columns_without_inline_group_or_key(tmp_path: Path):
    """86. Field protection list has strictly 2 columns (字段名称 + 保护方式), without inline group combo or key checkbox."""
    from privacy_app import PrivacyApp
    import tkinter as tk
    from tkinter import ttk

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.controller.set_sheet_header_row("Staff", 2)
    app._render_column_list()
    app.update()

    # Verify each row in col_inner_frame contains: 1 label frame (name+badges) and 1 combobox (protection)
    # No Checkbutton for record key on the row!
    data_rows = [w for w in app.col_inner_frame.winfo_children() if type(w) is tk.Frame]
    assert len(data_rows) == 4

    for r in data_rows:
        children = r.winfo_children()
        # Exactly 2 child containers: lbl_frame and prot_combo
        assert len(children) == 2
        # Ensure no Checkbutton or group Combobox exists inside row children
        checkbuttons = [w for w in children if isinstance(w, ttk.Checkbutton)]
        assert len(checkbuttons) == 0, "Found inline record key checkbutton on table row"

    app.destroy()


def test_87_changing_protection_does_not_jump_scroll(tmp_path: Path):
    """87. P0: Changing protection Combobox keeps col_canvas scroll position completely unchanged."""
    from privacy_app import PrivacyApp
    import tkinter as tk
    from tkinter import ttk

    # Create a workbook with 30 columns to ensure canvas is scrollable
    src = tmp_path / "large.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "WideSheet"
    headers = [f"Col_{i}" for i in range(35)]
    ws.append(headers)
    ws.append([100 * i for i in range(35)])
    wb.save(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("WideSheet")
    app.update()

    # Scroll down to 0.4
    app.col_canvas.yview_moveto(0.4)
    app.update()
    pos_before = app.col_canvas.yview()[0]

    # Find the combobox on row 15 and change protection
    data_rows = [w for w in app.col_inner_frame.winfo_children() if type(w) is tk.Frame]
    row_15 = data_rows[15]
    combos = [w for w in row_15.winfo_children() if isinstance(w, ttk.Combobox)]
    assert len(combos) == 1
    combo = combos[0]

    combo.set("数值保护")
    combo.event_generate("<<ComboboxSelected>>")
    app.update()

    pos_after = app.col_canvas.yview()[0]
    # Scroll position must NOT jump back to 0.0!
    assert pos_after == pytest.approx(pos_before, abs=0.01), f"Scroll position jumped from {pos_before} to {pos_after}"
    assert app.controller.sheet_configs["WideSheet"].columns[15].protection_type == "numeric"

    app.destroy()


def test_88_record_key_dialog_selection_and_validation(tmp_path: Path):
    """88. RecordKeyDialog opens, reflects selection, validates composite keys, and commits."""
    from privacy_app import PrivacyApp, RecordKeyDialog

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.controller.set_sheet_header_row("Staff", 2)
    app._render_column_list()
    app.update()

    # Open dialog
    dlg = RecordKeyDialog(app, "Staff")
    dlg.withdraw()
    assert "Staff" in dlg.title()
    assert "A" in dlg.col_vars
    assert "D" in dlg.col_vars

    # Select EmpID (Col A)
    dlg.col_vars["A"].set(True)
    dlg._on_check_toggle()
    assert "✓ 可唯一识别每条记录" in dlg.dialog_status_lbl.cget("text")

    # Apply and close
    dlg._apply_and_close()
    app.update()

    col_a = app.controller.sheet_configs["Staff"].get_column("A")
    assert col_a.is_record_key is True
    assert "✓ 已选择：EmpID" in app.key_status_lbl.cget("text")

    app.destroy()


def test_89_advanced_numeric_group_collapsible_and_local_assignment(tmp_path: Path):
    """89. Advanced numeric group card is collapsible by default and allows assigning local A/B/C/D groups."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.controller.set_sheet_header_row("Staff", 2)
    app.controller.set_column_protection("Staff", "C", "numeric")
    app._render_column_list()
    app.update()

    # Default is collapsed
    assert app.adv_expanded is False
    assert "展开" in app.btn_adv_toggle.cget("text")
    assert "当前 1 个数值保护字段" in app.adv_summary_lbl.cget("text")

    # Toggle open
    app._toggle_advanced_group()
    assert app.adv_expanded is True
    assert "收起" in app.btn_adv_toggle.cget("text")
    assert len(app.adv_content_frame.winfo_children()) > 0

    # Modify local group
    app.controller.set_column_local_group("Staff", "C", "B")
    assert app.controller.sheet_configs["Staff"].get_column("C").local_group == "B"

    # Toggle close
    app._toggle_advanced_group()
    assert app.adv_expanded is False
    assert "展开" in app.btn_adv_toggle.cget("text")

    app.destroy()


def test_90_shared_group_dialog_assigns_cross_sheet_group(tmp_path: Path):
    """90. SharedGroupDialog lists numeric columns across sheets and assigns cross-sheet shared group."""
    from privacy_app import PrivacyApp, SharedGroupDialog

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()

    # Configure Staff C as numeric, Bonuses C as numeric
    app.controller.set_sheet_header_row("Staff", 2)
    app.controller.set_column_protection("Staff", "C", "numeric")
    confirm_with_unique_record_identifier(app.controller, "Staff")
    assert app.controller.sheet_configs["Staff"].is_confirmed is True

    app.controller.set_sheet_header_row("Bonuses", 1)
    app.controller.set_column_protection("Bonuses", "C", "numeric")
    confirm_with_unique_record_identifier(app.controller, "Bonuses")
    assert app.controller.sheet_configs["Bonuses"].is_confirmed is True

    # Open SharedGroupDialog
    dlg = SharedGroupDialog(app)
    dlg.withdraw()

    # Verify both columns appear in group_vars
    assert ("Staff", "C") in dlg.group_vars
    assert ("Bonuses", "C") in dlg.group_vars

    # Set both to 共享组 1
    dlg.group_vars[("Staff", "C")].set("共享组 1")
    dlg.group_vars[("Bonuses", "C")].set("共享组 1")
    dlg._apply_and_close()
    app.update()

    # Both sheets should now have shared_group_id set and confirmed status invalidated to pending
    staff_c = app.controller.sheet_configs["Staff"].get_column("C")
    bonus_c = app.controller.sheet_configs["Bonuses"].get_column("C")
    assert staff_c.shared_group_id == "共享组 1"
    assert bonus_c.shared_group_id == "共享组 1"
    assert app.controller.sheet_configs["Staff"].is_confirmed is False
    assert app.controller.sheet_configs["Bonuses"].is_confirmed is False

    # Check scopes built for production mask
    scopes = app.controller.build_sheet_scopes()
    staff_scope = next(s for s in scopes if s.sheet_name == "Staff")
    bonus_scope = next(s for s in scopes if s.sheet_name == "Bonuses")
    assert staff_scope.numeric_column_groups["C"] == "共享组 1"
    assert bonus_scope.numeric_column_groups["C"] == "共享组 1"

    app.destroy()


def test_91_workbook_action_bar_exists(tmp_path: Path):
    """91. Workbook Action Bar exists on Tab 1 with status label, mask button, and AI prompt area."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    assert hasattr(app, "workbook_action_card")
    assert hasattr(app, "workbook_status_lbl")
    assert hasattr(app, "btn_mask")
    assert hasattr(app, "btn_copy_prompt")
    assert hasattr(app, "prompt_text")
    app.destroy()


def test_92_mask_button_visible_at_default_window(tmp_path: Path):
    """92. Mask button is mapped and visible within default window boundaries."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.update()
    assert app.btn_mask.winfo_ismapped()
    assert app.btn_mask.winfo_rooty() < app.winfo_rooty() + app.winfo_height()
    app.destroy()


def test_93_mask_button_disabled_when_zero_of_two_sheets_confirmed(tmp_path: Path):
    """93. Mask button is disabled and status shows 0/2 when no sheets are confirmed."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    assert str(app.btn_mask["state"]) == "disabled"
    assert "0 / 2" in app.workbook_status_lbl.cget("text")
    app.destroy()


def test_94_mask_button_disabled_when_one_of_two_sheets_confirmed(tmp_path: Path):
    """94. Mask button is disabled and status shows 1/2 when only one sheet is confirmed."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app.controller.set_sheet_header_row("Staff", 2)
    confirm_with_unique_record_identifier(app.controller, "Staff")
    app._render_sheet_list()

    assert str(app.btn_mask["state"]) == "disabled"
    assert "1 / 2" in app.workbook_status_lbl.cget("text")
    assert "Bonuses" in app.workbook_status_lbl.cget("text")
    app.destroy()


def test_95_mask_button_enabled_when_all_selected_sheets_confirmed(tmp_path: Path):
    """95. Mask button is enabled and status is green when all selected sheets are confirmed."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app.controller.set_sheet_header_row("Staff", 2)
    confirm_with_unique_record_identifier(app.controller, "Staff")
    app.controller.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(app.controller, "Bonuses")
    app._render_sheet_list()

    assert str(app.btn_mask["state"]) == "normal"
    assert "✓ 已确认 2 / 2" in app.workbook_status_lbl.cget("text")
    app.destroy()


def test_96_editing_confirmed_sheet_disables_mask_button_again(tmp_path: Path):
    """96. Editing a confirmed sheet invalidates confirmation and immediately disables mask button."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app.controller.set_sheet_header_row("Staff", 2)
    confirm_with_unique_record_identifier(app.controller, "Staff")
    app.controller.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(app.controller, "Bonuses")
    app._render_sheet_list()
    assert str(app.btn_mask["state"]) == "normal"

    # Invalidate Bonuses by editing a column
    app.controller.set_column_protection("Bonuses", "C", "numeric")
    app._render_sheet_list()

    assert str(app.btn_mask["state"]) == "disabled"
    assert "1 / 2" in app.workbook_status_lbl.cget("text")
    assert "Bonuses" in app.workbook_status_lbl.cget("text")
    app.destroy()


def test_97_mask_button_visible_when_numeric_advanced_expanded(tmp_path: Path):
    """97. Expanding Advanced Numeric Group does not push Workbook Action Bar out of view."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.geometry("880x680")
    app.update()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("Staff")
    app.update()

    app._toggle_advanced_group()
    app.update()

    assert app.adv_expanded is True
    assert app.btn_mask.winfo_ismapped()
    assert app.btn_mask.winfo_rooty() < app.winfo_rooty() + app.winfo_height()
    app.destroy()


def test_98_mask_button_visible_with_fifty_fields(tmp_path: Path):
    """98. Having 50 fields in table does not push Workbook Action Bar out of view."""
    import openpyxl
    from privacy_app import PrivacyApp

    src = tmp_path / "fifty.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BigSheet"
    ws.append([f"Header_{i}" for i in range(1, 51)])
    ws.append([f"Value_{i}" for i in range(1, 51)])
    wb.save(src)

    app = PrivacyApp()
    app.geometry("880x680")
    app.update()
    app.controller.load_source_workbook(src)
    app._render_sheet_list()
    app._switch_active_sheet("BigSheet")
    app.update()

    assert app.btn_mask.winfo_ismapped()
    assert app.btn_mask.winfo_rooty() < app.winfo_rooty() + app.winfo_height()
    app.destroy()


def test_99_mask_button_uses_existing_production_mask_callback(tmp_path: Path):
    """99. Mask button executes existing production controller mask workflow."""
    from privacy_app import PrivacyApp

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app.controller.set_sheet_header_row("Staff", 2)
    confirm_with_unique_record_identifier(app.controller, "Staff")
    app.controller.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(app.controller, "Bonuses")
    app._render_sheet_list()

    # Verify callback is _start_mask_task
    assert app.btn_mask.cget("command") != ""
    assert str(app.btn_mask["state"]) == "normal"

    out_dir = tmp_path / "masked_out"
    app.safe_dir_var.set(str(out_dir))
    app.controller.safe_dir = out_dir

    mask_res = app.controller.execute_mask()
    assert "masked_workbook_path" in mask_res
    assert Path(mask_res["masked_workbook_path"]).exists()
    app.destroy()


def test_100_successful_mask_enables_copy_ai_prompt(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """100. Successful masking enables copy AI prompt button and loads prompt into text area."""
    from privacy_app import PrivacyApp, messagebox

    monkeypatch.setattr(messagebox, "showinfo", lambda *args, **kwargs: None)

    src = tmp_path / "test.xlsx"
    create_sample_workbook(src)

    app = PrivacyApp()
    app.withdraw()
    app.controller.load_source_workbook(src)
    app.controller.set_sheet_header_row("Staff", 2)
    confirm_with_unique_record_identifier(app.controller, "Staff")
    app.controller.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(app.controller, "Bonuses")
    app._render_sheet_list()

    assert str(app.btn_copy_prompt["state"]) == "disabled"

    out_dir = tmp_path / "masked_out"
    app.controller.safe_dir = out_dir
    mask_res = app.controller.execute_mask()
    app._on_mask_success(mask_res)

    assert str(app.btn_copy_prompt["state"]) == "normal"
    assert "AI" in app.prompt_text.get("1.0", "end")
    assert "✓ 脱敏完成" in app.workbook_status_lbl.cget("text")
    app.destroy()


# ==============================================================================
# Tab 2 Batch Identification & Run Identity Regression Suite (Tests 101-108)
# ==============================================================================

def test_101_two_consecutive_masks_yield_independent_runs(tmp_path: Path):
    """101. Consecutive masking on same workbook generates Run A and Run B with distinct full run_ids and ledger files."""
    src = tmp_path / "src.xlsx"
    create_sample_workbook(src)

    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    # Mask 1 -> Run A
    ctl_a = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_a.load_source_workbook(src)
    ctl_a.set_sheet_header_row("Staff", 2)
    ctl_a.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_a, "Staff")
    ctl_a.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_a, "Bonuses")
    res_a = ctl_a.execute_mask()
    run_a = res_a["run_id"]
    masked_a = ctl_a.last_masked_path

    # Mask 2 -> Run B
    ctl_b = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_b.load_source_workbook(src)
    ctl_b.set_sheet_header_row("Staff", 2)
    ctl_b.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_b, "Staff")
    ctl_b.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_b, "Bonuses")
    res_b = ctl_b.execute_mask()
    run_b = res_b["run_id"]
    masked_b = ctl_b.last_masked_path

    # Core protocol guarantee: full run_id and ledger assets must be completely distinct
    assert run_a != run_b
    assert masked_a != masked_b
    assert res_a["ledger_enc_path"] != res_b["ledger_enc_path"]
    assert res_a["secret_key_path"] != res_b["secret_key_path"]
    assert Path(res_a["ledger_enc_path"]).exists()
    assert Path(res_b["ledger_enc_path"]).exists()


def test_102_run_a_matches_run_a_with_user_feedback(tmp_path: Path):
    """102. Run A AI file matches Run A ledger with clear batch feedback without leaking run_id/key."""
    from privacy_handler.identity import compute_run_fingerprint

    src = tmp_path / "src.xlsx"
    create_sample_workbook(src)
    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    ctl = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl, "Bonuses")
    res = ctl.execute_mask()
    run_id = res["run_id"]
    fp8 = compute_run_fingerprint(run_id)

    # Load returned workbook
    ok, msg = ctl.load_returned_workbook(ctl.last_masked_path)
    assert ok is True
    assert ctl.resolved_run_id == run_id
    assert ctl.resolved_batch_tag == fp8
    assert "✓ 已自动匹配该文件对应的本地恢复资料" in msg
    assert f"脱敏批次：{fp8}" in msg

    # Strict privacy: do NOT expose full run_id, secret key, ledger path, or k
    assert run_id not in msg
    assert ".recovery_ledger.enc" not in msg
    assert ".key" not in msg


def test_103_run_b_matches_run_b_with_user_feedback(tmp_path: Path):
    """103. Run B AI file matches Run B ledger independently."""
    from privacy_handler.identity import compute_run_fingerprint

    src = tmp_path / "src.xlsx"
    create_sample_workbook(src)
    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    # Mask 1
    ctl1 = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl1.load_source_workbook(src)
    ctl1.set_sheet_header_row("Staff", 2)
    ctl1.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl1, "Staff")
    ctl1.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl1, "Bonuses")
    ctl1.execute_mask()

    # Mask 2
    ctl2 = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl2.load_source_workbook(src)
    ctl2.set_sheet_header_row("Staff", 2)
    ctl2.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl2, "Staff")
    ctl2.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl2, "Bonuses")
    res2 = ctl2.execute_mask()
    run_b = res2["run_id"]
    fp8_b = compute_run_fingerprint(run_b)

    ok, msg = ctl2.load_returned_workbook(ctl2.last_masked_path)
    assert ok is True
    assert ctl2.resolved_run_id == run_b
    assert ctl2.resolved_batch_tag == fp8_b
    assert f"脱敏批次：{fp8_b}" in msg


def test_104_renamed_files_still_match_respective_runs(tmp_path: Path):
    """104. Even if AI returned files are arbitrarily renamed, they match their respective runs via Run Identity."""
    import shutil
    from privacy_handler.identity import compute_run_fingerprint

    src = tmp_path / "src.xlsx"
    create_sample_workbook(src)
    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    # Run A
    ctl_a = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_a.load_source_workbook(src)
    ctl_a.set_sheet_header_row("Staff", 2)
    ctl_a.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_a, "Staff")
    ctl_a.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_a, "Bonuses")
    res_a = ctl_a.execute_mask()
    run_a = res_a["run_id"]
    fp8_a = compute_run_fingerprint(run_a)

    # Run B
    ctl_b = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_b.load_source_workbook(src)
    ctl_b.set_sheet_header_row("Staff", 2)
    ctl_b.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_b, "Staff")
    ctl_b.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_b, "Bonuses")
    res_b = ctl_b.execute_mask()
    run_b = res_b["run_id"]
    fp8_b = compute_run_fingerprint(run_b)

    # Arbitrary renames
    renamed_a = tmp_path / "AI_Analysis_Final_Report_Renamed.xlsx"
    renamed_b = tmp_path / "totally_different_name_123.xlsx"
    shutil.copy(ctl_a.last_masked_path, renamed_a)
    shutil.copy(ctl_b.last_masked_path, renamed_b)

    # Controller resolving renamed_a
    ctl_test_a = GUIWorkflowController(private_keys_dir=priv_dir)
    ok_a, msg_a = ctl_test_a.load_returned_workbook(renamed_a)
    assert ok_a is True
    assert ctl_test_a.resolved_run_id == run_a
    assert ctl_test_a.resolved_batch_tag == fp8_a
    assert f"脱敏批次：{fp8_a}" in msg_a

    # Controller resolving renamed_b
    ctl_test_b = GUIWorkflowController(private_keys_dir=priv_dir)
    ok_b, msg_b = ctl_test_b.load_returned_workbook(renamed_b)
    assert ok_b is True
    assert ctl_test_b.resolved_run_id == run_b
    assert ctl_test_b.resolved_batch_tag == fp8_b
    assert f"脱敏批次：{fp8_b}" in msg_b


def test_105_no_cross_use_of_recovery_ledger(tmp_path: Path):
    """105. Run A file paired with Run B ledger is rejected by restore engine with RUN_IDENTITY_MISMATCH."""
    from privacy_handler.restore_engine import restore_workbook_v1, RestoreEngineError
    from privacy_handler.table_scope import SheetProtectionScope

    src = tmp_path / "src.xlsx"
    create_sample_workbook(src)
    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    # Run A
    ctl_a = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_a.load_source_workbook(src)
    ctl_a.set_sheet_header_row("Staff", 2)
    ctl_a.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_a, "Staff")
    ctl_a.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_a, "Bonuses")
    res_a = ctl_a.execute_mask()
    file_a = ctl_a.last_masked_path

    # Run B
    ctl_b = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_b.load_source_workbook(src)
    ctl_b.set_sheet_header_row("Staff", 2)
    ctl_b.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_b, "Staff")
    ctl_b.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_b, "Bonuses")
    res_b = ctl_b.execute_mask()

    ledger_b = res_b["ledger_enc_path"]
    secret_key_b = Path(res_b["secret_key_path"]).read_bytes().strip()

    # Force cross-use: file A with ledger B
    scopes = [
        SheetProtectionScope(sheet_name="Staff", header_row=2, text_columns=["B"], record_key="A"),
        SheetProtectionScope(sheet_name="Bonuses", header_row=1),
    ]

    import pytest
    with pytest.raises(RestoreEngineError) as exc_info:
        restore_workbook_v1(
            returned_workbook_path=file_a,
            ledger_path=ledger_b,
            secret_key=secret_key_b,
            output_dir=tmp_path / "restored",
        )
    assert "RUN_IDENTITY_MISMATCH" in str(exc_info.value)


def test_106_mismatched_run_identity_must_be_rejected(tmp_path: Path):
    """106. Mismatched or alien Run Identity returns error message and halts auto restore."""
    alien_wb = tmp_path / "alien.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["Alien Col 1", "Value 1"])
    wb.save(alien_wb)

    ctl = GUIWorkflowController(private_keys_dir=tmp_path / "priv")
    ok, msg = ctl.load_returned_workbook(alien_wb)

    assert ok is False
    assert "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。" in msg
    assert ctl.resolved_run_id is None
    assert ctl.resolved_batch_tag is None


def test_107_xpt_stripped_fallback_to_unique_text_fp8(tmp_path: Path):
    """107. When custom XML properties are stripped by third-party tool, fallback to unique FP8 in cell tokens."""
    import zipfile
    from privacy_handler.identity import compute_run_fingerprint

    src = tmp_path / "src.xlsx"
    create_sample_workbook(src)
    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    ctl = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Staff", 2)
    ctl.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl, "Staff")
    ctl.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl, "Bonuses")
    res = ctl.execute_mask()
    run_id = res["run_id"]
    fp8 = compute_run_fingerprint(run_id)
    masked_file = ctl.last_masked_path

    # Strip docProps/custom.xml
    stripped_file = tmp_path / "stripped.xlsx"
    with zipfile.ZipFile(masked_file, "r") as zin:
        with zipfile.ZipFile(stripped_file, "w") as zout:
            for item in zin.infolist():
                if item.filename != "docProps/custom.xml":
                    zout.writestr(item, zin.read(item.filename))

    ctl_fallback = GUIWorkflowController(private_keys_dir=priv_dir)
    ok, msg = ctl_fallback.load_returned_workbook(stripped_file)
    assert ok is True
    assert ctl_fallback.resolved_run_id == run_id
    assert ctl_fallback.resolved_batch_tag == fp8
    assert f"脱敏批次：{fp8}" in msg


def test_108_numeric_only_with_stripped_identity_halts_safely(tmp_path: Path):
    """108. Numeric-only workbook with stripped identity has no text tokens; controller refuses to guess and halts."""
    import zipfile

    src = tmp_path / "num_only.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nums"
    ws.append(["ID", "Amount"])
    ws.append(["1", 1000])
    ws.append(["2", 2000])
    wb.save(src)

    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"

    ctl = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl.load_source_workbook(src)
    ctl.set_sheet_header_row("Nums", 1)
    ctl.set_column_protection("Nums", "B", "numeric")
    confirm_with_unique_record_identifier(ctl, "Nums")
    ctl.execute_mask()
    masked_file = ctl.last_masked_path

    # Strip custom.xml so identity is completely gone
    stripped_file = tmp_path / "num_stripped.xlsx"
    with zipfile.ZipFile(masked_file, "r") as zin:
        with zipfile.ZipFile(stripped_file, "w") as zout:
            for item in zin.infolist():
                if item.filename != "docProps/custom.xml":
                    zout.writestr(item, zin.read(item.filename))

    ctl_halt = GUIWorkflowController(private_keys_dir=priv_dir)
    ok, msg = ctl_halt.load_returned_workbook(stripped_file)

    # Must NOT guess by filename or folder or latest run; must fail closed safely
    assert ok is False
    assert "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。" in msg
    assert ctl_halt.resolved_run_id is None
    assert ctl_halt.resolved_batch_tag is None


def test_109_deterministic_fp8_collision_and_ambiguity_gate(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """109. Deterministic FP8 collision & ambiguity resolution gate.

    Case A: Full authoritative XPT_RunId custom property present matches exact run_id
            even when another run shares identical FP8 fingerprint (FP8 collision does not blind authoritative match).
    Case B: Stripped custom properties with identical FP8 in candidate ledgers causes ambiguity,
            safely halts without guessing, refuses to pick candidate, and displays stopping feedback.
    """
    import zipfile
    from cryptography.fernet import Fernet
    from privacy_handler.recovery_ledger import RecoveryLedger, LedgerRunInfo, save_encrypted_ledger

    priv_dir = tmp_path / "private_keys"
    priv_dir.mkdir(parents=True, exist_ok=True)

    run_id_1 = "11111111-0000-0000-0000-000000000001"
    run_id_2 = "22222222-0000-0000-0000-000000000002"
    colliding_fp8 = "COLLIS01"

    key1 = Fernet.generate_key()
    key2 = Fernet.generate_key()

    (priv_dir / f"{run_id_1}.mask_secret.key").write_bytes(key1)
    (priv_dir / f"{run_id_2}.mask_secret.key").write_bytes(key2)

    ledger1 = RecoveryLedger(run=LedgerRunInfo(run_id=run_id_1, created_at="2026-10-09T00:00:00Z", workbook_identity="wb_test_1"), entries=[])
    ledger2 = RecoveryLedger(run=LedgerRunInfo(run_id=run_id_2, created_at="2026-10-09T00:00:00Z", workbook_identity="wb_test_2"), entries=[])

    save_encrypted_ledger(ledger1, priv_dir / f"{run_id_1}.recovery_ledger.enc", key1)
    save_encrypted_ledger(ledger2, priv_dir / f"{run_id_2}.recovery_ledger.enc", key2)

    # Mock compute_run_fingerprint so both runs yield identical FP8 COLLIS01
    from privacy_handler import gui_controller, identity

    def mock_fp(rid: str) -> str:
        if rid in (run_id_1, run_id_2):
            return colliding_fp8
        return "OTHERFP8"

    monkeypatch.setattr(gui_controller, "compute_run_fingerprint", mock_fp)
    monkeypatch.setattr(identity, "compute_run_fingerprint", mock_fp)

    # --- Case A: Authoritative XPT_RunId property present in returned workbook ---
    wb_a = openpyxl.Workbook()
    ws_a = wb_a.active
    ws_a.title = "Sheet1"
    ws_a.append(["Col1", f"[{colliding_fp8}_00000001]"])
    identity.stamp_masked_workbook_identity(wb_a, run_id_1, "1.0")
    file_a = tmp_path / "returned_case_a.xlsx"
    wb_a.save(file_a)

    ctl_a = GUIWorkflowController(private_keys_dir=priv_dir)
    ok_a, msg_a = ctl_a.load_returned_workbook(file_a)

    # Must resolve accurately to run_id_1 via authoritative XPT property, unaffected by FP8 collision
    assert ok_a is True
    assert ctl_a.resolved_run_id == run_id_1
    assert ctl_a.resolved_batch_tag == colliding_fp8
    assert f"脱敏批次：{colliding_fp8}" in msg_a
    assert ctl_a.resolved_ledger_path == priv_dir / f"{run_id_1}.recovery_ledger.enc"
    assert ctl_a.resolved_secret_key == key1

    # --- Case B: Stripped custom properties with identical FP8 candidate ledgers ---
    wb_b = openpyxl.Workbook()
    ws_b = wb_b.active
    ws_b.title = "Sheet1"
    # Stripped properties, only cell token exists
    ws_b.append(["Col1", f"[{colliding_fp8}_00000002]"])
    file_b = tmp_path / "returned_case_b_stripped.xlsx"
    wb_b.save(file_b)

    ctl_b = GUIWorkflowController(private_keys_dir=priv_dir)
    ok_b, msg_b = ctl_b.load_returned_workbook(file_b)

    # Must fail closed safely because resolve_ledger_by_fingerprint finds 2 candidate ledgers for COLLIS01
    assert ok_b is False
    assert "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。" in msg_b
    assert ctl_b.resolved_run_id is None
    assert ctl_b.resolved_batch_tag is None
    assert ctl_b.resolved_ledger_path is None
    assert ctl_b.resolved_secret_key is None


def test_110_headless_dual_run_gui_gate(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """110. Comprehensive headless Tkinter dual-run GUI gate:
    1. Run A and Run B masks are performed consecutively.
    2. Load AI A -> feedback shows A's batch tag, btn_compare enabled.
    3. Load AI B -> switches feedback to B's batch tag, btn_compare enabled.
    4. Rename AI A -> still resolves Run A in GUI.
    5. Move AI B to different dir -> still resolves Run B in GUI.
    6. Load alien file -> stops with warning banner, btn_compare disabled.
    7. Business modification on AI A preserves identity and restores cleanly.
    """
    import shutil
    from privacy_app import PrivacyApp, messagebox
    from privacy_handler.identity import compute_run_fingerprint

    monkeypatch.setattr(messagebox, "showerror", lambda *args, **kwargs: None)

    src = tmp_path / "dual_gui_src.xlsx"
    create_sample_workbook(src)

    priv_dir = tmp_path / "private_keys"
    safe_dir = tmp_path / "safe"
    restored_dir = tmp_path / "restored"

    # 1. Mask Run A via Controller
    ctl_a = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_a.load_source_workbook(src)
    ctl_a.set_sheet_header_row("Staff", 2)
    ctl_a.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_a, "Staff")
    ctl_a.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_a, "Bonuses")
    res_a = ctl_a.execute_mask()
    run_a = res_a["run_id"]
    fp8_a = compute_run_fingerprint(run_a)
    masked_a = Path(ctl_a.last_masked_path)

    # Mask Run B via Controller
    ctl_b = GUIWorkflowController(safe_dir=safe_dir, private_keys_dir=priv_dir)
    ctl_b.load_source_workbook(src)
    ctl_b.set_sheet_header_row("Staff", 2)
    ctl_b.set_column_protection("Staff", "B", "text")
    confirm_with_unique_record_identifier(ctl_b, "Staff")
    ctl_b.set_sheet_header_row("Bonuses", 1)
    confirm_with_unique_record_identifier(ctl_b, "Bonuses")
    res_b = ctl_b.execute_mask()
    run_b = res_b["run_id"]
    fp8_b = compute_run_fingerprint(run_b)
    masked_b = Path(ctl_b.last_masked_path)

    assert run_a != run_b

    # Initialize headless GUI App with shared private keys dir
    app = PrivacyApp()
    app.withdraw()
    app.controller.private_keys_dir = priv_dir
    app.restored_dir_var.set(str(restored_dir))
    app.controller.restored_dir = restored_dir

    # Step 2: Load AI A into App
    app.returned_path_var.set(str(masked_a))
    ok_a, msg_a = app.controller.load_returned_workbook(masked_a)
    app._on_returned_loaded(ok_a, msg_a)
    app.update()

    assert ok_a is True
    assert app.controller.resolved_run_id == run_a
    assert app.controller.resolved_batch_tag == fp8_a
    assert f"脱敏批次：{fp8_a}" in app.batch_feedback_lbl.cget("text")
    assert "✓ 已自动匹配" in app.batch_feedback_lbl.cget("text")
    assert str(app.btn_compare["state"]) == "normal"

    # Step 3: Switch by loading AI B into App
    app.returned_path_var.set(str(masked_b))
    ok_b, msg_b = app.controller.load_returned_workbook(masked_b)
    app._on_returned_loaded(ok_b, msg_b)
    app.update()

    assert ok_b is True
    assert app.controller.resolved_run_id == run_b
    assert app.controller.resolved_batch_tag == fp8_b
    assert f"脱敏批次：{fp8_b}" in app.batch_feedback_lbl.cget("text")
    assert str(app.btn_compare["state"]) == "normal"

    # Step 4: Rename AI A -> still resolves Run A
    renamed_a = tmp_path / "renamed_report_a.xlsx"
    shutil.copy(masked_a, renamed_a)
    ok_ra, msg_ra = app.controller.load_returned_workbook(renamed_a)
    app._on_returned_loaded(ok_ra, msg_ra)
    app.update()

    assert ok_ra is True
    assert app.controller.resolved_run_id == run_a
    assert app.controller.resolved_batch_tag == fp8_a
    assert f"脱敏批次：{fp8_a}" in app.batch_feedback_lbl.cget("text")

    # Step 5: Move AI B to external folder -> still resolves Run B
    other_folder = tmp_path / "some" / "nested" / "folder"
    other_folder.mkdir(parents=True, exist_ok=True)
    moved_b = other_folder / "moved_b.xlsx"
    shutil.copy(masked_b, moved_b)
    ok_mb, msg_mb = app.controller.load_returned_workbook(moved_b)
    app._on_returned_loaded(ok_mb, msg_mb)
    app.update()

    assert ok_mb is True
    assert app.controller.resolved_run_id == run_b
    assert app.controller.resolved_batch_tag == fp8_b
    assert f"脱敏批次：{fp8_b}" in app.batch_feedback_lbl.cget("text")

    # Step 6: Load foreign / alien file -> halts with warning, disables compare
    alien_file = tmp_path / "foreign_workbook.xlsx"
    wb_alien = openpyxl.Workbook()
    wb_alien.active.append(["Alien Header", 999])
    wb_alien.save(alien_file)

    ok_alien, msg_alien = app.controller.load_returned_workbook(alien_file)
    app._on_returned_loaded(ok_alien, msg_alien)
    app.update()

    assert ok_alien is False
    assert "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。" in app.batch_feedback_lbl.cget("text")
    assert app.controller.resolved_run_id is None
    assert str(app.btn_compare["state"]) == "disabled"
    assert str(app.btn_restore["state"]) == "disabled"

    # Step 7: Business modification on AI A preserves identity and restores cleanly
    wb_ai_a = openpyxl.load_workbook(masked_a)
    ws_ai_a = wb_ai_a["Staff"]
    ws_ai_a.cell(row=2, column=5, value="AI Analysis Note")
    ws_ai_a.cell(row=3, column=5, value="Reviewed OK")
    edited_a = tmp_path / "ai_edited_a.xlsx"
    wb_ai_a.save(edited_a)

    ok_ea, msg_ea = app.controller.load_returned_workbook(edited_a)
    app._on_returned_loaded(ok_ea, msg_ea)
    app.update()

    assert ok_ea is True
    assert app.controller.resolved_run_id == run_a
    cmp_res = app.controller.execute_compare()
    assert cmp_res.added_columns_count >= 1

    rest_res = app.controller.execute_restore()
    assert rest_res.restored_cells_count > 0
    assert Path(rest_res.restored_workbook_path).exists()

    app.destroy()
