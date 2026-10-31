"""Unit and regression tests for Checkpoint 3.5: TableScope, Column Selection, and OutputPaths."""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys
import tempfile
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl.styles import PatternFill

import privacy_handler.excel_privacy_tool as tool
from privacy_handler.table_scope import (
    ColumnDescriptor,
    OutputPaths,
    TableScope,
    TableScopeError,
    extract_table_scope,
    iter_table_scope_rows,
    resolve_safe_output_path,
    validate_composite_key,
)


# ==============================================================================
# Section 20 Tests: TableScope, Header Row & Column Selection (Tests 1 - 25)
# ==============================================================================

def test_default_header_row_1():
    """1. Default header_row=1 behavior matches traditional tables."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Age", "Department"])
    ws.append(["Alice", 30, "Engineering"])
    ws.append(["Bob", 25, "Design"])

    scope = extract_table_scope(ws, header_row=1)
    assert scope.header_row == 1
    assert scope.data_start_row == 2
    assert scope.data_end_row == 3
    assert len(scope.selected_columns) == 3
    headers = [col.header for col in scope.selected_columns]
    assert headers == ["Name", "Age", "Department"]


def test_header_row_5_custom_title():
    """2. header_row=5 correctly ignores rows 1-4 title/notes and starts data at row 6."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append(["ACME Corporation Monthly Financials"])  # Row 1
    ws.append(["Period: 2026-Q1"])                      # Row 2
    ws.append([])                                        # Row 3 (empty)
    ws.append(["Confidential - Internal Only"])          # Row 4
    ws.append(["Employee ID", "Name", "Salary"])         # Row 5 (Header)
    ws.append(["E001", "Alice", 10000])                  # Row 6
    ws.append(["E002", "Bob", 12000])                    # Row 7

    scope = extract_table_scope(ws, header_row=5)
    assert scope.header_row == 5
    assert scope.data_start_row == 6
    assert scope.data_end_row == 7
    assert len(scope.selected_columns) == 3
    assert [col.header for col in scope.selected_columns] == ["Employee ID", "Name", "Salary"]


def test_data_start_row_explicit():
    """3. Explicit data_start_row=7 when row 6 is an extra comment row."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Employee ID", "Name", "Salary"])         # Row 1 (Header)
    ws.append(["(Note: in USD)", "", ""])                # Row 2 (Subheader/Note)
    ws.append(["E001", "Alice", 10000])                  # Row 3 (Data)
    ws.append(["E002", "Bob", 12000])                    # Row 4 (Data)

    scope = extract_table_scope(ws, header_row=1, data_start_row=3)
    assert scope.header_row == 1
    assert scope.data_start_row == 3
    assert scope.data_end_row == 4
    rows = iter_table_scope_rows(ws, scope)
    assert len(rows) == 2
    assert rows[0]["Employee ID"] == "E001"
    assert rows[1]["Employee ID"] == "E002"


def test_invalid_header_row_less_than_1():
    """4. header_row < 1 raises safe TableScopeError."""
    wb = openpyxl.Workbook()
    ws = wb.active
    with pytest.raises(TableScopeError, match="header_row must be >= 1"):
        extract_table_scope(ws, header_row=0)


def test_invalid_data_start_row_le_header_row():
    """5. data_start_row <= header_row raises safe TableScopeError."""
    wb = openpyxl.Workbook()
    ws = wb.active
    with pytest.raises(TableScopeError, match="strictly greater than header_row"):
        extract_table_scope(ws, header_row=3, data_start_row=3)
    with pytest.raises(TableScopeError, match="strictly greater than header_row"):
        extract_table_scope(ws, header_row=3, data_start_row=2)


def test_empty_header_cells_displayed_as_placeholder():
    """6. Empty header cells are rendered as [空表头] without fabricating column names."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="ID")
    ws.cell(row=1, column=2, value="")        # Empty string
    ws.cell(row=1, column=3, value=None)      # None
    ws.cell(row=1, column=4, value="   ")     # Whitespace

    scope = extract_table_scope(ws, header_row=1)
    headers = [c.header for c in scope.selected_columns]
    assert headers == ["ID", "[空表头]", "[空表头]", "[空表头]"]
    # Check internal raw_header semantics
    assert scope.selected_columns[0].raw_header == "ID"
    assert scope.selected_columns[0].is_blank_header is False
    assert scope.selected_columns[1].raw_header == ""
    assert scope.selected_columns[1].is_blank_header is True
    assert scope.selected_columns[2].raw_header is None
    assert scope.selected_columns[2].is_blank_header is True
    assert scope.selected_columns[3].raw_header == "   "
    assert scope.selected_columns[3].is_blank_header is True


def test_literal_kong_biaotou_distinguished_from_blank():
    """6b. A column whose literal header is '[空表头]' is distinguished from an actually blank header."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="[空表头]")  # Literal header named '[空表头]'
    ws.cell(row=1, column=2, value=None)        # Truly blank cell

    scope = extract_table_scope(ws, header_row=1)
    col1 = scope.selected_columns[0]
    col2 = scope.selected_columns[1]

    # Both display '[空表头]' for human/UI
    assert col1.display_header == "[空表头]"
    assert col2.display_header == "[空表头]"

    # Internally, raw_header preserves true origin
    assert col1.raw_header == "[空表头]"
    assert col1.is_blank_header is False

    assert col2.raw_header is None
    assert col2.is_blank_header is True


def test_duplicate_headers_distinguished_by_letter():
    """7. Duplicate header names are cleanly distinguished by column letters."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="ID")
    ws.cell(row=1, column=2, value="金额")
    ws.cell(row=1, column=3, value="金额")
    ws.cell(row=2, column=1, value="E01")
    ws.cell(row=2, column=2, value=100)
    ws.cell(row=2, column=3, value=200)

    scope = extract_table_scope(ws, header_row=1)
    cols = scope.selected_columns
    assert cols[1].column_letter == "B"
    assert cols[1].header == "金额"
    assert cols[2].column_letter == "C"
    assert cols[2].header == "金额"

    rows = iter_table_scope_rows(ws, scope)
    assert len(rows) == 1
    assert rows[0]["B"] == 100
    assert rows[0]["C"] == 200
    assert rows[0]["B｜金额"] == 100
    assert rows[0]["C｜金额"] == 200


def test_hidden_columns_flagged_not_dropped():
    """8. Hidden columns have is_hidden=True and are not silently dropped."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A_col", "B_col", "C_col"])
    ws.append([1, 2, 3])
    ws.column_dimensions["B"].hidden = True

    scope = extract_table_scope(ws, header_row=1)
    assert len(scope.selected_columns) == 3
    assert scope.selected_columns[0].is_hidden is False
    assert scope.selected_columns[1].is_hidden is True
    assert scope.selected_columns[2].is_hidden is False


def test_column_letter_and_index_1_indexed():
    """9. ColumnDescriptor column_letter and column_index are strictly 1-indexed."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", "B", "C"])

    scope = extract_table_scope(ws, header_row=1)
    assert scope.selected_columns[0].column_letter == "A"
    assert scope.selected_columns[0].column_index == 1
    assert scope.selected_columns[2].column_letter == "C"
    assert scope.selected_columns[2].column_index == 3


def test_column_selection_subset():
    """10. Selecting a subset of column letters filters properly."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Department", "Salary"])
    ws.append(["1", "Alice", "IT", 5000])

    scope = extract_table_scope(ws, header_row=1, selected_column_letters=["A", "D"])
    assert len(scope.selected_columns) == 2
    assert [c.column_letter for c in scope.selected_columns] == ["A", "D"]
    assert [c.header for c in scope.selected_columns] == ["ID", "Salary"]


def test_trailing_empty_formatted_rows_trimmed():
    """11. Trailing formatted empty rows are trimmed from data_end_row."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name"])
    ws.append(["1", "Alice"])
    ws.append(["2", "Bob"])

    # Rows 4 to 8: empty values but cell has fill style
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for r in range(4, 9):
        cell = ws.cell(row=r, column=1)
        cell.fill = fill
        cell.value = None

    scope = extract_table_scope(ws, header_row=1)
    assert scope.data_end_row == 3  # trimmed from 8 down to 3


def test_empty_sheet_no_columns():
    """12. Completely empty worksheet produces empty selected_columns and None data_end_row."""
    wb = openpyxl.Workbook()
    ws = wb.active
    scope = extract_table_scope(ws, header_row=1)
    assert scope.selected_columns == []
    assert scope.data_end_row is None


def test_sheet_with_only_header_no_data():
    """13. Worksheet with only header row produces data_end_row=None."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Score"])
    scope = extract_table_scope(ws, header_row=1)
    assert len(scope.selected_columns) == 3
    assert scope.data_end_row is None
    assert iter_table_scope_rows(ws, scope) == []


def test_recommended_protection_pii_keywords():
    """14. Recommended protection flag correctly identifies PII keywords."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["员工姓名", "身份证号", "手机号码", "内部编号", "部门"])
    scope = extract_table_scope(ws, header_row=1)
    rec = {c.header: c.recommended_protection for c in scope.selected_columns}
    assert rec["员工姓名"] is True
    assert rec["身份证号"] is True
    assert rec["手机号码"] is True
    assert rec["内部编号"] is False
    assert rec["部门"] is False


def test_numeric_candidate_detection():
    """15. is_numeric_candidate accurately identifies columns containing numbers."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Salary", "Bonus", "Active"])
    ws.append(["E1", "Alice", 10000, 500.5, True])
    ws.append(["E2", "Bob", 12000, 300.0, False])

    scope = extract_table_scope(ws, header_row=1)
    cand = {c.header: c.is_numeric_candidate for c in scope.selected_columns}
    assert cand["ID"] is False
    assert cand["Name"] is False
    assert cand["Salary"] is True
    assert cand["Bonus"] is True
    assert cand["Active"] is False  # booleans are not numeric candidates


def test_composite_key_two_unique_columns():
    """16. Composite key on two columns with unique pairs is valid."""
    rows = [
        {"Dept": "Engineering", "EmployeeID": "101"},
        {"Dept": "Engineering", "EmployeeID": "102"},
        {"Dept": "Marketing", "EmployeeID": "101"},
        {"Dept": "Marketing", "EmployeeID": "102"},
    ]
    res = validate_composite_key(rows, ["Dept", "EmployeeID"])
    assert res["is_valid"] is True
    assert res["total_count"] == 4
    assert res["unique_count"] == 4
    assert res["null_count"] == 0


def test_composite_key_duplicate_pairs():
    """17. Composite key with duplicate pair fails validation."""
    rows = [
        {"Dept": "Engineering", "EmployeeID": "101"},
        {"Dept": "Engineering", "EmployeeID": "101"},  # duplicate
    ]
    res = validate_composite_key(rows, ["Dept", "EmployeeID"])
    assert res["is_valid"] is False
    assert "composite key has duplicates" in res["reason"]


def test_composite_key_with_null_fails():
    """18. Composite key where any row has null/empty in a component fails."""
    rows = [
        {"Dept": "Engineering", "EmployeeID": "101"},
        {"Dept": "", "EmployeeID": "102"},  # empty Dept
    ]
    res = validate_composite_key(rows, ["Dept", "EmployeeID"])
    assert res["is_valid"] is False
    assert res["null_count"] == 1


def test_composite_key_with_formula_fails():
    """19. Composite key where a row has a formula value fails."""
    rows = [
        {"Dept": "Engineering", "EmployeeID": "101"},
        {"Dept": "=VLOOKUP(A1, B:C, 2, 0)", "EmployeeID": "102"},
    ]
    res = validate_composite_key(rows, ["Dept", "EmployeeID"])
    assert res["is_valid"] is False
    assert res["formula_count"] == 1


def test_composite_key_empty_fields_fails():
    """20. Validating composite key with empty key_fields list fails."""
    rows = [{"ID": "1"}]
    res = validate_composite_key(rows, [])
    assert res["is_valid"] is False
    assert "key_fields cannot be empty" in res["reason"]


def test_composite_key_empty_rows_fails():
    """21. Validating composite key on 0 rows fails."""
    res = validate_composite_key([], ["ID"])
    assert res["is_valid"] is False
    assert res["total_count"] == 0


def test_composite_key_missing_field_raises():
    """22. Specifying a field not in rows raises safe TableScopeError."""
    rows = [{"ID": "1"}]
    with pytest.raises(TableScopeError, match="Field 'NonExistent' not present"):
        validate_composite_key(rows, ["NonExistent"])


def test_composite_key_preserves_typed_identity():
    """23. Composite key does not conflate int(1) and str('1')."""
    rows = [
        {"Type": "A", "Code": 1},
        {"Type": "A", "Code": "1"},
    ]
    res = validate_composite_key(rows, ["Type", "Code"])
    assert res["is_valid"] is True
    assert res["unique_count"] == 2


def test_table_scope_serialization_roundtrip():
    """24. TableScope to_dict and from_dict preserve all properties."""
    scope = TableScope(
        sheet_name="Payroll",
        header_row=5,
        data_start_row=6,
        data_end_row=100,
        selected_columns=[
            ColumnDescriptor("A", 1, "ID", "ID", False, True, False),
            ColumnDescriptor("B", 2, "Salary", "Salary", True, False, True),
        ],
    )
    d = scope.to_dict()
    restored = TableScope.from_dict(d)
    assert restored.sheet_name == "Payroll"
    assert restored.header_row == 5
    assert restored.data_start_row == 6
    assert restored.data_end_row == 100
    assert len(restored.selected_columns) == 2
    assert restored.selected_columns[0].header == "ID"
    assert restored.selected_columns[1].is_hidden is True
    assert restored.selected_columns[1].is_numeric_candidate is True


def test_table_scope_module_is_integrity_locked():
    """25. table_scope.py is included in LOCKED_FILES."""
    table_scope_file = Path(__file__).resolve().parent.parent / "privacy_handler" / "table_scope.py"
    assert table_scope_file in tool.LOCKED_FILES


# ==============================================================================
# Section 18.10 Tests: OutputPaths Decoupling & Collision Safety (Tests 26 - 33)
# ==============================================================================

def test_output_paths_decoupled(tmp_path: Path):
    """26. OutputPaths keeps masked, restored, and private directories separate."""
    m_dir = tmp_path / "custom_masked"
    r_dir = tmp_path / "custom_restored"
    p_dir = tmp_path / "secure_ledger"
    paths = OutputPaths(m_dir, r_dir, p_dir)
    assert paths.masked_output_dir != paths.private_data_dir
    assert paths.restored_output_dir != paths.private_data_dir


def test_output_paths_serialization_roundtrip(tmp_path: Path):
    """27. OutputPaths serialization to_dict / from_dict roundtrip."""
    paths = OutputPaths(tmp_path / "m", tmp_path / "r", tmp_path / "p")
    d = paths.to_dict()
    restored = OutputPaths.from_dict(d)
    assert restored.masked_output_dir == paths.masked_output_dir.resolve()
    assert restored.restored_output_dir == paths.restored_output_dir.resolve()
    assert restored.private_data_dir == paths.private_data_dir.resolve()


def test_safe_output_path_does_not_overwrite_source(tmp_path: Path):
    """28. Safe output path never collides with or overwrites source file."""
    source_file = tmp_path / "data.xlsx"
    source_file.write_text("dummy")

    # If base dir is same as source dir, resolved path must not equal source
    out = resolve_safe_output_path(tmp_path, source_file, suffix=".xlsx")
    assert out != source_file
    assert out.name == "data_1.xlsx"


def test_safe_output_path_collision_avoidance(tmp_path: Path):
    """29. If target file already exists, appends _1, _2 increment safely."""
    source = tmp_path / "sales.xlsx"
    source.write_text("source")

    out_dir = tmp_path / "exports"
    out_dir.mkdir()

    # Pre-create files
    (out_dir / "sales.masked.xlsx").write_text("existing1")
    (out_dir / "sales.masked_1.xlsx").write_text("existing2")

    target = resolve_safe_output_path(out_dir, source, suffix=".masked.xlsx")
    assert target.name == "sales.masked_2.xlsx"
    assert not target.exists()


def test_safe_output_path_creates_missing_directory(tmp_path: Path):
    """30. Creates non-existent target output directory automatically."""
    target_dir = tmp_path / "deep" / "nested" / "dir"
    source = tmp_path / "test.xlsx"
    assert not target_dir.exists()

    out = resolve_safe_output_path(target_dir, source)
    assert target_dir.exists()
    assert out.parent == target_dir


def test_safe_output_path_chinese_characters_and_spaces(tmp_path: Path):
    """31. Supports directory and file paths with Chinese characters and spaces."""
    chinese_dir = tmp_path / "财务 报表 输出"
    source = tmp_path / "2026年 薪酬明细 表.xlsx"
    source.write_text("test")

    out = resolve_safe_output_path(chinese_dir, source)
    assert "财务 报表 输出" in str(out)
    assert "2026年 薪酬明细 表.masked.xlsx" in out.name


def test_safe_output_path_unwritable_directory_raises(tmp_path: Path):
    """32. Unwritable output directory raises safe TableScopeError."""
    ro_dir = tmp_path / "readonly_dir"
    ro_dir.mkdir()
    os.chmod(ro_dir, 0o444)
    source = tmp_path / "data.xlsx"

    try:
        # Check if root or user can still write; if so, test is skipped
        if not os.access(ro_dir, os.W_OK):
            with pytest.raises(TableScopeError, match="is not writable"):
                resolve_safe_output_path(ro_dir, source)
    finally:
        os.chmod(ro_dir, 0o755)


def test_private_ledger_dir_never_comingles_with_masked_export(tmp_path: Path):
    """33. Invariant check: private ledger directory is not colocated with masked dir."""
    user_export_dir = tmp_path / "Desktop" / "SharedToAI"
    app_private_dir = tmp_path / "AppData" / "SecureLedger"

    paths = OutputPaths(
        masked_output_dir=user_export_dir,
        restored_output_dir=user_export_dir / "restored",
        private_data_dir=app_private_dir,
    )
    # Ensure distinct trees
    assert app_private_dir not in paths.masked_output_dir.parents
    assert paths.masked_output_dir != paths.private_data_dir
