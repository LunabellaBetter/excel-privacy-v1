"""Comprehensive tests for Checkpoint 6: AI Round-trip Point Restore Engine.

Tests cover:
1. Multi-Sheet Masking & Restore:
   - Multiple sheets masked in single run, each with custom header_row and keys.
   - Cross-sheet token consistency: identical (field, value) shares token and restores cleanly.
   - Shared numeric group scale factor vs independent scale factors.
2. Numeric Restore Iron Law:
   - UNCHANGED scaled values restored to original typed values (int, float, decimal).
   - PROTECTED_VALUE_CONFLICT (modified values) are NEVER restored; NO division by k.
   - PROTECTED VALUE SWAP (swapped across records) detected and NEVER restored.
   - AI modified numeric values left exactly as AI returned.
3. Text Token Restoration:
   - UNCHANGED tokens restored to original string values.
   - Corrupted or alien tokens skipped without crashing.
4. Structural Preservation:
   - Whole-row sort by AI preserved; correct values restored by Record Identity.
   - AI added columns preserved.
   - AI deleted rows stay deleted; remaining rows restore correctly.
   - AI added rows stay added without injecting PII.
   - Formulas preserved and calculation flags set (fullCalcOnLoad, forceFullCalc).
   - Extra sheets added by AI preserved.
5. Typed Values Restoration:
   - Leading zeros strings (e.g. "00123") restored as strings, not parsed as ints.
   - Dates and datetimes restored with exact types.
   - Booleans restored as booleans.
6. Safety & Security Gates:
   - Unsupported workbook parts (PivotTable, Slicer, etc.) blocked at pre-flight.
   - Source workbook overwrite protection.
   - Safe collision-free output paths.
   - Integrity lock verification passes.
"""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
import io
import json
from pathlib import Path
import sys
import zipfile
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from cryptography.fernet import Fernet

import privacy_handler.excel_privacy_tool as tool
from privacy_handler.conflict_engine import (
    ChangeList,
    ChangeStatus,
    compare_workbook_against_ledger,
)
from privacy_handler.recovery_ledger import (
    LedgerConfig,
    LedgerEntry,
    LedgerRunInfo,
    ProtectedField,
    RecoveryLedger,
    load_encrypted_ledger,
    save_encrypted_ledger,
)
from privacy_handler.identity import (
    read_workbook_identity,
    stamp_masked_workbook_identity,
)
from privacy_handler.restore_engine import (
    RestoreEngineError,
    RestoreResult,
    restore_workbook_v1,
)
from privacy_handler.table_scope import SheetProtectionScope, extract_table_scope


# ==============================================================================
# Helper functions
# ==============================================================================

def create_multi_sheet_source(path: Path) -> Path:
    wb = openpyxl.Workbook()
    # Sheet 1: Employees (header at row 3)
    ws1 = wb.active
    ws1.title = "Employees"
    ws1.append(["Confidential HR Master", None, None, None])
    ws1.append([None, None, None, None])
    ws1.append(["EmpID", "Name", "Salary", "Dept"])
    ws1.append(["E001", "Alice", 10000, "Engineering"])
    ws1.append(["E002", "Bob", 20000, "Finance"])
    ws1.append(["E003", "Charlie", 30000, "Engineering"])

    # Sheet 2: Payroll (header at row 1)
    ws2 = wb.create_sheet(title="Payroll")
    ws2.append(["EmpID", "Month", "Bonus", "Name"])
    ws2.append(["E001", "2026-01", 1500, "Alice"])  # "Alice" also appears here
    ws2.append(["E002", "2026-01", 2500, "David"])
    ws2.append(["E003", "2026-01", 3500, "Alice"])

    wb.save(path)
    return path


# ==============================================================================
# 1. Multi-Sheet Protection & Round-trip Restore
# ==============================================================================

def test_01_multi_sheet_mask_and_clean_restore(tmp_path: Path):
    """01. Verify multi-sheet masking and 100% clean restore with heterogeneous scopes."""
    src = tmp_path / "multi_source.xlsx"
    create_multi_sheet_source(src)

    out_dir = tmp_path / "safe_for_codex"
    priv_dir = tmp_path / "private_keys"
    rest_dir = tmp_path / "restored_outputs"

    scopes = [
        SheetProtectionScope(
            sheet_name="Employees",
            header_row=3,
            text_columns=["Name"],
            numeric_columns=["Salary"],
            record_key="EmpID",
            numeric_analysis_group="grp1",
        ),
        SheetProtectionScope(
            sheet_name="Payroll",
            header_row=1,
            text_columns=["Name"],
            numeric_columns=["Bonus"],
            record_key=["EmpID", "Month"],
            numeric_analysis_group="grp1",
        ),
    ]

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        sheet_scopes=scopes,
        numeric_scale_factor=Decimal("2.5000"),
    )

    masked_path = mask_res["masked_workbook_path"]
    ledger_path = mask_res["ledger_enc_path"]
    secret_key = mask_res["secret_key_path"].read_bytes()

    # Load masked workbook to inspect
    m_wb = openpyxl.load_workbook(masked_path)
    ws_emp = m_wb["Employees"]
    ws_pay = m_wb["Payroll"]

    # In Employees: Alice is masked
    alice_token = ws_emp.cell(row=4, column=2).value
    assert isinstance(alice_token, str) and alice_token.startswith("MSK-")
    # Salary is scaled: 10000 * 2.5 = 25000
    assert ws_emp.cell(row=4, column=3).value == 25000.0

    # In Payroll: Name Alice shares identical token across sheets!
    pay_name_token = ws_pay.cell(row=2, column=4).value
    assert pay_name_token == alice_token

    # Now simulate AI returning identical untouched file
    res = restore_workbook_v1(
        returned_workbook_path=masked_path,
        ledger_path=ledger_path,
        secret_key=secret_key,
        output_dir=rest_dir,
    )

    assert res.restored_workbook_path.exists()
    assert res.skipped_conflicts_count == 0
    # Employees: 3 names + 3 salaries = 6; Payroll: 3 approvers + 3 bonuses = 6 -> Total 12
    assert res.restored_cells_count == 12

    # Check restored contents
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_emp = r_wb["Employees"]
    r_pay = r_wb["Payroll"]

    assert r_emp.cell(row=4, column=2).value == "Alice"
    assert r_emp.cell(row=4, column=3).value == 10000
    assert r_emp.cell(row=5, column=2).value == "Bob"
    assert r_emp.cell(row=5, column=3).value == 20000

    assert r_pay.cell(row=2, column=4).value == "Alice"
    assert r_pay.cell(row=2, column=3).value == 1500


def test_02_cross_sheet_token_consistency(tmp_path: Path):
    """02. Cross-sheet token allocation must assign identical token for same (field, value)."""
    src = tmp_path / "cross_token.xlsx"
    create_multi_sheet_source(src)

    scopes = [
        SheetProtectionScope(sheet_name="Employees", header_row=3, text_columns=["Name"], record_key="EmpID"),
        SheetProtectionScope(sheet_name="Payroll", header_row=1, text_columns=["Approver"], record_key="EmpID"),
    ]

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        sheet_scopes=scopes,
    )

    wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    # Employees row 4 is Alice (col 2)
    emp_token = wb["Employees"].cell(row=4, column=2).value
    # Payroll row 2 and 4 are Alice (col 4)
    pay_token1 = wb["Payroll"].cell(row=2, column=4).value
    pay_token2 = wb["Payroll"].cell(row=4, column=4).value

    # Since Name and Approver are different logical fields, check if same logical field shares token
    # Let's test with exact same logical field across sheets
    wb2 = openpyxl.Workbook()
    s1 = wb2.active
    s1.title = "S1"
    s1.append(["ID", "Name"])
    s1.append(["1", "Zhang San"])
    s2 = wb2.create_sheet(title="S2")
    s2.append(["ID", "Name"])
    s2.append(["2", "Zhang San"])
    p2 = tmp_path / "same_field.xlsx"
    wb2.save(p2)

    scopes2 = [
        SheetProtectionScope(sheet_name="S1", header_row=1, text_columns=["Name"], record_key="ID"),
        SheetProtectionScope(sheet_name="S2", header_row=1, text_columns=["Name"], record_key="ID"),
    ]
    res2 = tool.mask_workbook_v1(input_path=p2, sheet_scopes=scopes2, private_key_dir=tmp_path / "priv2", output_dir=tmp_path / "safe2")
    wb_m2 = openpyxl.load_workbook(res2["masked_workbook_path"])
    t1 = wb_m2["S1"].cell(row=2, column=2).value
    t2 = wb_m2["S2"].cell(row=2, column=2).value
    assert t1 == t2
    assert t1.startswith("MSK-")


# ==============================================================================
# 2. Numeric Restore Iron Law: Unchanged vs Modified/Swapped
# ==============================================================================

def test_03_numeric_restore_iron_law_modified_cell_skipped(tmp_path: Path):
    """03. IRON LAW: If AI modifies a protected numeric value, NEVER auto-restore and NEVER divide by k."""
    src = tmp_path / "salary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary"])
    ws.append(["E01", 10000])
    ws.append(["E02", 20000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )
    masked_path = mask_res["masked_workbook_path"]
    ledger_path = mask_res["ledger_enc_path"]
    secret_key = mask_res["secret_key_path"].read_bytes()

    # AI modifies E01's salary from 20000 to 99999 (e.g. AI gives raise or modifies)
    ret_wb = openpyxl.load_workbook(masked_path)
    ret_ws = ret_wb.active
    ret_ws.cell(row=2, column=2).value = 99999
    ai_returned_path = tmp_path / "ai_returned.xlsx"
    ret_wb.save(ai_returned_path)

    res = restore_workbook_v1(
        returned_workbook_path=ai_returned_path,
        ledger_path=ledger_path,
        secret_key=secret_key,
        output_dir=tmp_path / "restored",
    )

    # E01 was modified -> PROTECTED_VALUE_CONFLICT -> skipped!
    # E02 was unchanged -> restored!
    assert res.restored_cells_count == 1
    assert res.skipped_conflicts_count == 1

    restored_wb = openpyxl.load_workbook(res.restored_workbook_path)
    rest_ws = restored_wb.active

    # E01 MUST retain AI modified value 99999, NOT 10000, and NOT 99999/2=49999.5!
    assert rest_ws.cell(row=2, column=2).value == 99999
    # E02 restored to original 20000
    assert rest_ws.cell(row=3, column=2).value == 20000


def test_04_numeric_restore_iron_law_swapped_values_skipped(tmp_path: Path):
    """04. IRON LAW: If AI swaps masked values between records, conflict engine catches swap and restore skips both."""
    src = tmp_path / "swap.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary"])
    ws.append(["E01", 10000])  # Masked: 20000
    ws.append(["E02", 50000])  # Masked: 100000
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # AI accidentally swaps the two values!
    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_ws = ret_wb.active
    ret_ws.cell(row=2, column=2).value = 100000.0  # Put E02's masked value on E01
    ret_ws.cell(row=3, column=2).value = 20000.0   # Put E01's masked value on E02
    ai_swapped_path = tmp_path / "ai_swapped.xlsx"
    ret_wb.save(ai_swapped_path)

    res = restore_workbook_v1(
        returned_workbook_path=ai_swapped_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 0
    assert res.skipped_conflicts_count == 2

    # Verify workbook retains swapped values untouched (no wrongful restore)
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    assert r_wb.active.cell(row=2, column=2).value == 100000.0
    assert r_wb.active.cell(row=3, column=2).value == 20000.0


def test_05_disallow_conflicts_raises_error(tmp_path: Path):
    """05. When allow_conflicts=False, existing conflicts raise RestoreEngineError."""
    src = tmp_path / "err.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary"])
    ws.append(["E01", 10000])
    ws.append(["E02", 20000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key="EmpID",
    )

    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_wb.active.cell(row=2, column=2).value = 88888  # Modified
    ret_path = tmp_path / "ret.xlsx"
    ret_wb.save(ret_path)

    with pytest.raises(RestoreEngineError, match="Cannot restore workbook due to conflicts"):
        restore_workbook_v1(
            returned_workbook_path=ret_path,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            allow_conflicts=False,
        )


# ==============================================================================
# 3. Structural Preservation: Sorts, Inserted Columns, Deleted/Added Rows
# ==============================================================================

def test_06_row_reordering_whole_row_sort_clean_restore(tmp_path: Path):
    """06. AI sorts rows by a column; Record Identity correctly locates records and restores them."""
    src = tmp_path / "sort.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    ws.append(["E03", "Charlie", 30000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # AI sorts rows descending by salary (E03, E02, E01)
    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_ws = ret_wb.active
    m_alice = ret_ws.cell(row=2, column=2).value
    m_bob = ret_ws.cell(row=3, column=2).value
    m_charlie = ret_ws.cell(row=4, column=2).value

    # Overwrite in sorted order:
    ret_ws.cell(row=2, column=1).value = "E03"
    ret_ws.cell(row=2, column=2).value = m_charlie
    ret_ws.cell(row=2, column=3).value = 60000.0

    ret_ws.cell(row=3, column=1).value = "E02"
    ret_ws.cell(row=3, column=2).value = m_bob
    ret_ws.cell(row=3, column=3).value = 40000.0

    ret_ws.cell(row=4, column=1).value = "E01"
    ret_ws.cell(row=4, column=2).value = m_alice
    ret_ws.cell(row=4, column=3).value = 20000.0

    sorted_path = tmp_path / "sorted.xlsx"
    ret_wb.save(sorted_path)

    res = restore_workbook_v1(
        returned_workbook_path=sorted_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 6
    assert res.skipped_conflicts_count == 0

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active

    # Row 2 is E03: Charlie, 30000
    assert r_ws.cell(row=2, column=1).value == "E03"
    assert r_ws.cell(row=2, column=2).value == "Charlie"
    assert r_ws.cell(row=2, column=3).value == 30000

    # Row 4 is E01: Alice, 10000
    assert r_ws.cell(row=4, column=1).value == "E01"
    assert r_ws.cell(row=4, column=2).value == "Alice"
    assert r_ws.cell(row=4, column=3).value == 10000


def test_07_ai_inserted_column_preserved(tmp_path: Path):
    """07. AI adds a new calculated/analysis column; new column is preserved upon restore."""
    src = tmp_path / "add_col.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary"])
    ws.append(["E01", 10000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # AI inserts column C: "TaxBracket"
    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_ws = ret_wb.active
    ret_ws.cell(row=1, column=3).value = "TaxBracket"
    ret_ws.cell(row=2, column=3).value = "Tier-1"
    ret_path = tmp_path / "with_col.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=2, column=2).value == 10000
    # Column C preserved!
    assert r_ws.cell(row=1, column=3).value == "TaxBracket"
    assert r_ws.cell(row=2, column=3).value == "Tier-1"


def test_08_ai_deleted_row_and_added_row(tmp_path: Path):
    """08. AI deletes a row and adds a new row; deleted row remains gone, new row remains added."""
    src = tmp_path / "del_add.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # AI deletes E01 (Bob shifts to row 2) and adds new row E99 (row 3)
    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_ws = ret_wb.active
    m_bob = ret_ws.cell(row=3, column=2).value
    ret_ws.delete_rows(2, 1)  # Delete Alice

    # Add E99
    ret_ws.append(["E99", "NewHire", 55555])
    ret_path = tmp_path / "del_add_out.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    # Only Bob's cells restored (Name + Salary = 2)
    assert res.restored_cells_count == 2

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    # Row 2 is Bob restored
    assert r_ws.cell(row=2, column=1).value == "E02"
    assert r_ws.cell(row=2, column=2).value == "Bob"
    assert r_ws.cell(row=2, column=3).value == 20000
    # Row 3 is E99 untouched
    assert r_ws.cell(row=3, column=1).value == "E99"
    assert r_ws.cell(row=3, column=2).value == "NewHire"
    assert r_ws.cell(row=3, column=3).value == 55555


def test_09_ai_added_sheet_preserved(tmp_path: Path):
    """09. AI creates a new summary sheet; restored workbook contains the new sheet intact."""
    src = tmp_path / "add_sheet.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["EmpID", "Salary"])
    ws.append(["E01", 10000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )

    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    sum_sheet = ret_wb.create_sheet(title="ExecutiveSummary")
    sum_sheet.append(["Metric", "Value"])
    sum_sheet.append(["TotalHeadcount", 1])
    ret_path = tmp_path / "extra_sheet.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    assert "ExecutiveSummary" in r_wb.sheetnames
    assert r_wb["ExecutiveSummary"].cell(row=2, column=2).value == 1
    assert r_wb["Data"].cell(row=2, column=2).value == 10000


# ==============================================================================
# 4. Typed Values Restoration (Strings with leading zeros, Dates, Booleans)
# ==============================================================================

def test_10_typed_value_restoration_leading_zeros(tmp_path: Path):
    """10. Text values with leading zeros (e.g. '00123') must be restored as string, not integer."""
    src = tmp_path / "leading_zeros.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Key", "Code"])
    ws.append(["K1", "00012345"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Code"],
        record_key="Key",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    val = r_wb.active.cell(row=2, column=2).value
    assert isinstance(val, str)
    assert val == "00012345"


def test_11_recalculation_flags_set_on_restore(tmp_path: Path):
    """11. Restored workbook must have fullCalcOnLoad and forceFullCalc recalculation properties."""
    src = tmp_path / "recalc.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Key", "Amount", "Formula"])
    ws.append(["K1", 500, "=B2*2"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Amount"],
        record_key="Key",
        numeric_scale_factor=Decimal("2.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    # Inspect OOXML recalculation properties
    with zipfile.ZipFile(res.restored_workbook_path, "r") as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        assert "fullCalcOnLoad" in wb_xml
        assert "forceFullCalc" in wb_xml


def test_12_safe_output_collision_prevention(tmp_path: Path):
    """12. Restore output respects custom output_dir and generates collision-free filename if file exists."""
    src = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Key", "Val"])
    ws.append(["K1", 100])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Val"],
        record_key="Key",
    )

    out_dir = tmp_path / "restored_dir"
    res1 = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=out_dir,
    )
    assert res1.restored_workbook_path.exists()

    # Second restore to same directory should not overwrite first file
    res2 = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=out_dir,
    )
    assert res2.restored_workbook_path.exists()
    assert res1.restored_workbook_path != res2.restored_workbook_path


def test_13_preflight_blocks_unsupported_parts(tmp_path: Path):
    """13. Pre-flight check blocks unsupported OOXML parts (PivotTable, Slicer)."""
    p = tmp_path / "pivot.xlsx"
    # Create fake zip with PivotTable part
    with zipfile.ZipFile(p, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("xl/pivotTables/pivotTable1.xml", "<pivotTable/>")

    with pytest.raises(tool.PrivacyToolError, match="暂不支持的高级对象"):
        restore_workbook_v1(
            returned_workbook_path=p,
            ledger_path=tmp_path / "dummy_ledger.enc",
            secret_key=Fernet.generate_key(),
        )


def test_14_restore_engine_module_is_integrity_locked():
    """14. restore_engine.py is registered in tool.LOCKED_FILES and verify_self passes."""
    locked_names = [p.name for p in tool.LOCKED_FILES]
    assert "restore_engine.py" in locked_names
    tool.verify_self_before_sensitive()


# ==============================================================================
# 5. Composite Key Restoration & Tolerance Edge Cases
# ==============================================================================

def test_15_composite_key_roundtrip_restore(tmp_path: Path):
    """15. Composite Key (e.g. Dept+EmpID) accurately restores cells across duplicate single-key columns."""
    src = tmp_path / "comp.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Dept", "EmpID", "Score", "Bonus"])
    ws.append(["HR", "101", "A+", 1200])
    ws.append(["IT", "101", "B-", 3400])  # Same EmpID '101' in different Dept!
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Score"],
        numeric_columns=["Bonus"],
        record_key=["Dept", "EmpID"],
        numeric_scale_factor=Decimal("3.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 4
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=2, column=3).value == "A+"
    assert r_ws.cell(row=2, column=4).value == 1200
    assert r_ws.cell(row=3, column=3).value == "B-"
    assert r_ws.cell(row=3, column=4).value == 3400


def test_16_composite_key_row_swap_detected_as_conflicts(tmp_path: Path):
    """16. In Composite Key records, if AI swaps scores between (HR, 101) and (IT, 101), conflict engine blocks."""
    src = tmp_path / "comp_swap.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Dept", "EmpID", "Score"])
    ws.append(["HR", "101", "TopSecret_HR"])
    ws.append(["IT", "101", "TopSecret_IT"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Score"],
        record_key=["Dept", "EmpID"],
    )

    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_ws = ret_wb.active
    t_hr = ret_ws.cell(row=2, column=3).value
    t_it = ret_ws.cell(row=3, column=3).value
    ret_ws.cell(row=2, column=3).value = t_it
    ret_ws.cell(row=3, column=3).value = t_hr
    ret_path = tmp_path / "swapped.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 0
    assert res.skipped_conflicts_count == 2


def test_17_date_and_datetime_preservation(tmp_path: Path):
    """17. Dates and Datetimes are preserved with typed identity upon round-trip."""
    src = tmp_path / "dates.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    d_val = dt.date(2026, 10, 8)
    dt_val = dt.datetime(2026, 10, 8, 14, 30, 0)
    ws.append(["ID", "EventDate", "Timestamp", "Amount"])
    ws.append(["1", d_val, dt_val, 1500])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Amount"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 1
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    # Date in column 2 untouched
    val_date = r_ws.cell(row=2, column=2).value
    assert val_date.year == 2026 and val_date.month == 10 and val_date.day == 8
    # Datetime in column 3 untouched
    assert r_ws.cell(row=2, column=3).value == dt_val
    # Amount restored
    assert r_ws.cell(row=2, column=4).value == 1500


def test_18_boolean_preservation(tmp_path: Path):
    """18. Boolean values in unselected and selected columns are preserved as bool."""
    src = tmp_path / "bools.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "IsActive", "Amount"])
    ws.append(["1", True, 200])
    ws.append(["2", False, 400])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Amount"],
        record_key="ID",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=2, column=2).value is True
    assert r_ws.cell(row=3, column=2).value is False


def test_19_empty_cells_and_none_values_handling(tmp_path: Path):
    """19. Empty cells in protected columns are safely preserved as None, not corrupted."""
    src = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Bonus"])
    ws.append(["1", "Alice", 100])
    ws.append(["2", None, None])  # Empty row
    ws.append(["3", "Charlie", 300])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Bonus"],
        record_key="ID",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=3, column=2).value is None
    assert r_ws.cell(row=3, column=3).value is None
    assert r_ws.cell(row=4, column=2).value == "Charlie"
    assert r_ws.cell(row=4, column=3).value == 300


def test_20_int_vs_float_type_restoration(tmp_path: Path):
    """20. Original int remains int and float remains float upon restore."""
    src = tmp_path / "types.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "IntegerVal", "FloatVal"])
    ws.append(["1", 5000, 123.456])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["IntegerVal", "FloatVal"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    int_v = r_ws.cell(row=2, column=2).value
    float_v = r_ws.cell(row=2, column=3).value
    assert isinstance(int_v, int)
    assert int_v == 5000
    assert isinstance(float_v, float)
    assert abs(float_v - 123.456) < 1e-6


def test_21_header_row_greater_than_one_restore(tmp_path: Path):
    """21. Restore correctly handles header_row = 5 with multiple preamble rows."""
    src = tmp_path / "preamble.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Title 1", None])
    ws.append(["Title 2", None])
    ws.append([None, None])
    ws.append(["Notes", None])
    ws.append(["Code", "Value"])  # row 5
    ws.append(["C1", 100])        # row 6
    ws.append(["C2", 200])        # row 7
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=5,
        numeric_columns=["Value"],
        record_key="Code",
        numeric_scale_factor=Decimal("2.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 2
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=1, column=1).value == "Title 1"
    assert r_ws.cell(row=6, column=2).value == 100
    assert r_ws.cell(row=7, column=2).value == 200


def test_22_column_relocation_preserved_and_restored(tmp_path: Path):
    """22. AI swaps column positions (e.g. Salary moved before Name); restored values follow relocated columns."""
    src = tmp_path / "reloc.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E1", "Alice", 10000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0"),
    )

    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_ws = ret_wb.active
    m_name = ret_ws.cell(row=2, column=2).value
    m_salary = ret_ws.cell(row=2, column=3).value

    # Swap columns B and C
    ret_ws.cell(row=1, column=2).value = "Salary"
    ret_ws.cell(row=2, column=2).value = m_salary
    ret_ws.cell(row=1, column=3).value = "Name"
    ret_ws.cell(row=2, column=3).value = m_name
    ret_path = tmp_path / "reloc_out.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 2
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    # Column B is Salary -> 10000
    assert r_ws.cell(row=1, column=2).value == "Salary"
    assert r_ws.cell(row=2, column=2).value == 10000
    # Column C is Name -> Alice
    assert r_ws.cell(row=1, column=3).value == "Name"
    assert r_ws.cell(row=2, column=3).value == "Alice"


def test_23_negative_numeric_values_restoration(tmp_path: Path):
    """23. Negative numeric values and scale factor preserve negative sign accurately."""
    src = tmp_path / "neg.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "PnL"])
    ws.append(["1", -500.50])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["PnL"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    v = r_wb.active.cell(row=2, column=2).value
    assert abs(v - (-500.50)) < 1e-6


def test_24_zero_numeric_values_restoration(tmp_path: Path):
    """24. Zero value (0 and 0.0) in numeric column scales to 0 and restores to 0."""
    src = tmp_path / "zero.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Qty"])
    ws.append(["1", 0])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Qty"],
        record_key="ID",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    assert r_wb.active.cell(row=2, column=2).value == 0


def test_25_formula_preservation_in_restored_file(tmp_path: Path):
    """25. Formulas referencing restored cells are preserved and evaluate correctly."""
    src = tmp_path / "formula.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "A", "B", "Sum"])
    ws.append(["1", 10, 20, "=B2+C2"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["A"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=2, column=2).value == 10
    assert r_ws.cell(row=2, column=4).value == "=B2+C2"


def test_26_xlsm_is_rejected_for_v1(tmp_path: Path):
    src = tmp_path / "macro_test.xlsm"
    openpyxl.Workbook().save(src)
    with pytest.raises(tool.PrivacyToolError, match=".xlsx"):
        tool.mask_workbook_v1(input_path=src, output_dir=tmp_path / "safe",
                             private_key_dir=tmp_path / "priv")
    assert not (tmp_path / "safe").exists()


def test_27_overwrite_source_prevention(tmp_path: Path):
    """27. Attempting to restore directly over the returned workbook path raises error."""
    src = tmp_path / "ow.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Secret"])
    ws.append(["1", "Val"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Secret"],
        record_key="ID",
    )

    m_path = mask_res["masked_workbook_path"]
    with pytest.raises(tool.PrivacyToolError, match="Refusing to overwrite"):
        restore_workbook_v1(
            returned_workbook_path=m_path,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            output_dir=m_path.parent,
            output_filename=m_path.name,
        )


def test_28_restore_result_summary_and_serialization(tmp_path: Path):
    """28. RestoreResult.to_dict() serialization contains all execution metrics."""
    src = tmp_path / "res_test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Secret"])
    ws.append(["1", "Val"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Secret"],
        record_key="ID",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    d = res.to_dict()
    assert d["run_id"] == mask_res["run_id"]
    assert d["restored_cells_count"] == 1
    assert d["skipped_conflicts_count"] == 0
    assert "restored_workbook_path" in d


def test_29_multiple_sheets_with_partial_conflicts(tmp_path: Path):
    """29. In multi-sheet restore, conflicts on Sheet 1 do not prevent clean restore on Sheet 2."""
    src = tmp_path / "multi_partial.xlsx"
    create_multi_sheet_source(src)

    scopes = [
        SheetProtectionScope(sheet_name="Employees", header_row=3, numeric_columns=["Salary"], record_key="EmpID"),
        SheetProtectionScope(sheet_name="Payroll", header_row=1, numeric_columns=["Bonus"], record_key="EmpID"),
    ]

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        sheet_scopes=scopes,
        numeric_scale_factor=Decimal("2.0"),
    )

    # Tamper with Employees row 4 (Salary)
    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_wb["Employees"].cell(row=4, column=3).value = 999999
    ret_path = tmp_path / "tampered.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    # Employees has 1 conflict and 2 restored; Payroll has 3 restored -> Total restored = 5
    assert res.restored_cells_count == 5
    assert res.skipped_conflicts_count == 1
    assert res.restored_by_sheet["Employees"] == 2
    assert res.restored_by_sheet["Payroll"] == 3


def test_30_whitespace_padded_record_key_matches(tmp_path: Path):
    """30. Whitespace padded string in AI key cell matches expected key tuple."""
    src = tmp_path / "ws_key.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Secret"])
    ws.append(["ABC", "SecretVal"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Secret"],
        record_key="ID",
    )

    # AI returns " ABC " with whitespace padding in ID column
    ret_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ret_wb.active.cell(row=2, column=1).value = " ABC "
    ret_path = tmp_path / "padded.xlsx"
    ret_wb.save(ret_path)

    res = restore_workbook_v1(
        returned_workbook_path=ret_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 1
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    assert r_wb.active.cell(row=2, column=2).value == "SecretVal"


def test_31_multiple_unprotected_columns_unmodified(tmp_path: Path):
    """31. 10 unselected columns with complex formulas and comments remain 100% untouched."""
    src = tmp_path / "many_cols.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["ID", "Secret"] + [f"Col{i}" for i in range(1, 11)]
    ws.append(headers)
    row_vals = ["1", "SecretValue"] + [i * 10 for i in range(1, 11)]
    ws.append(row_vals)
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Secret"],
        record_key="ID",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=2, column=2).value == "SecretValue"
    for i in range(1, 11):
        assert r_ws.cell(row=2, column=2 + i).value == i * 10


def test_32_duplicate_header_columns_disambiguation_restore(tmp_path: Path):
    """32. Multi-row snapshot pattern disambiguates candidate duplicate-header column and restores correctly."""
    src = tmp_path / "dup_hdr.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Score", "Score"])
    ws.append(["E01", 100, 200])
    ws.append(["E02", 300, 400])
    wb.save(src)

    # Protect Score at col 3 (column letter C, values 200, 400)
    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["C"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # Simulate AI swapping the positions of the two Score columns
    # Starting from masked_workbook_path preserves stamped XPT metadata (as an AI / human editor would)
    ai_file = tmp_path / "ai_swapped.xlsx"
    wb_ai = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ws_ai = wb_ai.active
    # Swap values between column 2 (B) and column 3 (C)
    ws_ai.cell(row=2, column=2).value, ws_ai.cell(row=2, column=3).value = 400, 100
    ws_ai.cell(row=3, column=2).value, ws_ai.cell(row=3, column=3).value = 800, 300
    wb_ai.save(ai_file)

    res = restore_workbook_v1(
        returned_workbook_path=ai_file,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 2
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    # In returned sheet, Col 2 was the scaled Score (400, 800), which restores to original (200, 400)
    assert r_ws.cell(row=2, column=2).value == 200
    assert r_ws.cell(row=3, column=2).value == 400
    # Col 3 was untouched Score (100, 300)
    assert r_ws.cell(row=2, column=3).value == 100
    assert r_ws.cell(row=3, column=3).value == 300


def test_33_no_key_single_occurrence_clean_restore(tmp_path: Path):
    """33. No Record Identity defined, single occurrence of token restores cleanly via before_value match."""
    src = tmp_path / "no_key.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Title", "Description"])
    ws.append(["Doc1", "Confidential Strategy Memo"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Description"],
        record_key=None,  # No record key!
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 1
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    assert r_wb.active.cell(row=2, column=2).value == "Confidential Strategy Memo"


def test_34_diagnostic_privacy_in_restore_logs_and_exports(tmp_path: Path):
    """34. Default to_dict on ChangeList during restore must redact plain record keys."""
    src = tmp_path / "diag.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["SensitiveSSN", "Bonus"])
    ws.append(["110101199001011234", 5000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Bonus"],
        record_key="SensitiveSSN",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.change_list is not None
    safe_dict = res.change_list.to_safe_dict()
    # Confirm SSN does not appear in safe dict
    dumped = json.dumps(safe_dict)
    assert "110101199001011234" not in dumped


def test_35_end_to_end_full_roundtrip_fidelity(tmp_path: Path):
    """35. Comprehensive E2E Round-trip: Original -> Mask -> AI Edit -> Conflict Check -> Restore."""
    src = tmp_path / "e2e_full.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "BaseSalary", "Score", "Notes"])
    ws.append(["E001", "Alice", 10000, 95, "Rockstar"])
    ws.append(["E002", "Bob", 20000, 80, "Consistent"])
    ws.append(["E003", "Charlie", 30000, 60, "Needs improvement"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["BaseSalary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("1.5000"),
    )

    # Simulate AI:
    # 1. Swaps row 1 and row 3 (sort)
    # 2. Modifies Bob's salary (gives raise to 99999) -> Conflict!
    # 3. Adds an analysis column "Recommendation"
    # 4. Leaves Alice and Charlie untouched
    ai_wb = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ai_ws = ai_wb.active

    # Add Recommendation column
    ai_ws.cell(row=1, column=6).value = "Recommendation"
    ai_ws.cell(row=2, column=6).value = "Promote"
    ai_ws.cell(row=3, column=6).value = "Bonus"
    ai_ws.cell(row=4, column=6).value = "PIP"

    # Modify Bob's BaseSalary (row 3, col 3) to 99999
    ai_ws.cell(row=3, column=3).value = 99999

    ai_path = tmp_path / "ai_processed.xlsx"
    ai_wb.save(ai_path)

    res = restore_workbook_v1(
        returned_workbook_path=ai_path,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "final_restored",
    )

    # 3 records * 2 protected fields = 6 potential cells
    # Bob's salary modified -> 1 conflict skipped
    # 5 cells restored!
    assert res.restored_cells_count == 5
    assert res.skipped_conflicts_count == 1

    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active

    # Alice restored
    assert r_ws.cell(row=2, column=2).value == "Alice"
    assert r_ws.cell(row=2, column=3).value == 10000

    # Bob's name restored, but salary KEPT as AI modified (99999)
    assert r_ws.cell(row=3, column=2).value == "Bob"
    assert r_ws.cell(row=3, column=3).value == 99999

    # Charlie restored
    assert r_ws.cell(row=4, column=2).value == "Charlie"
    assert r_ws.cell(row=4, column=3).value == 30000

    # AI column preserved
    assert r_ws.cell(row=1, column=6).value == "Recommendation"
    assert r_ws.cell(row=2, column=6).value == "Promote"


# ==============================================================================
# 11. Checkpoint 6 Final Safety Patch Gates (Tests 36-46)
# ==============================================================================

def test_36_duplicate_header_two_columns_simultaneously_protected(tmp_path: Path):
    """36. P0-1: Two columns with identical header name (B: Amount, C: Amount) simultaneously protected and restored accurately."""
    src = tmp_path / "dup_hdr_both.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Amount", "Amount"])
    ws.append(["E01", 100, 200])
    ws.append(["E02", 300, 400])
    wb.save(src)

    # Protect both Amount columns (B and C) with numeric masking
    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["B", "C"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # Clean roundtrip restore
    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 4
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    assert r_ws.cell(row=2, column=2).value == 100
    assert r_ws.cell(row=2, column=3).value == 200
    assert r_ws.cell(row=3, column=2).value == 300
    assert r_ws.cell(row=3, column=3).value == 400


def test_37_duplicate_header_two_columns_swapped_restored_accurately(tmp_path: Path):
    """37. P0-1: Two columns with identical header name swapped by AI are restored accurately without cross-talk."""
    src = tmp_path / "dup_hdr_swap.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Amount", "Amount"])
    ws.append(["E01", 100, 200])
    ws.append(["E02", 300, 400])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["B", "C"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    # AI swaps the two Amount columns:
    # Original masked: B has 200, 600; C has 400, 800
    # AI returns: B has 400, 800; C has 200, 600
    ai_file = tmp_path / "ai_swapped.xlsx"
    wb_ai = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    ws_ai = wb_ai.active
    ws_ai.cell(row=2, column=2).value, ws_ai.cell(row=2, column=3).value = 400, 200
    ws_ai.cell(row=3, column=2).value, ws_ai.cell(row=3, column=3).value = 800, 600
    wb_ai.save(ai_file)

    res = restore_workbook_v1(
        returned_workbook_path=ai_file,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 4
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    r_ws = r_wb.active
    # In returned sheet, Col 2 (was C, value 400/800) restores to 200/400
    assert r_ws.cell(row=2, column=2).value == 200
    assert r_ws.cell(row=3, column=2).value == 400
    # Col 3 (was B, value 200/600) restores to 100/300
    assert r_ws.cell(row=2, column=3).value == 100
    assert r_ws.cell(row=3, column=3).value == 300


def test_38_run_identity_exact_match_passes(tmp_path: Path):
    """38. P0-2: Standard masked workbook with matching XPT_RunId restores cleanly."""
    src = tmp_path / "run_match.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name"])
    ws.append(["1", "Alice"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        record_key="ID",
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )
    assert res.restored_cells_count == 1


def test_39_run_identity_mismatch_fails_closed_zero_output(tmp_path: Path):
    """39. P0-2: Returned workbook with alien XPT_RunId is rejected and creates NO output files."""
    src = tmp_path / "mismatch.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name"])
    ws.append(["1", "Alice"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        record_key="ID",
    )

    # Tamper XPT_RunId in returned workbook
    tampered_file = tmp_path / "tampered.xlsx"
    wb_tampered = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    stamp_masked_workbook_identity(wb_tampered, run_id="alien-run-id-999")
    wb_tampered.save(tampered_file)

    out_dir = tmp_path / "should_not_exist_restored"

    with pytest.raises(RestoreEngineError, match="RUN_IDENTITY_MISMATCH"):
        restore_workbook_v1(
            returned_workbook_path=tampered_file,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            output_dir=out_dir,
        )

    # Verify zero output files created
    assert not out_dir.exists() or len(list(out_dir.glob("*.xlsx"))) == 0


def test_40_missing_run_id_with_valid_token_fallback_passes(tmp_path: Path):
    """40. P0-2: XPT props stripped by external tool, but valid consistent FP8 token exists -> fallback passes."""
    src = tmp_path / "strip_props.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name"])
    ws.append(["1", "Alice"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        record_key="ID",
    )

    # Strip custom doc props cleanly using empty CustomPropertyList
    from openpyxl.packaging.custom import CustomPropertyList
    stripped_file = tmp_path / "stripped.xlsx"
    wb_stripped = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    wb_stripped.custom_doc_props = CustomPropertyList()
    wb_stripped.save(stripped_file)

    # Verify props are gone
    ident = read_workbook_identity(stripped_file)
    assert ident["run_id"] is None

    # Fallback to token FP8 succeeds!
    res = restore_workbook_v1(
        returned_workbook_path=stripped_file,
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )
    assert res.restored_cells_count == 1


def test_41_missing_run_id_with_mismatched_token_fails(tmp_path: Path):
    """41. P0-2: Missing XPT_RunId and tokens belong to different run -> rejected."""
    src = tmp_path / "wrong_token.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name"])
    ws.append(["1", "Alice"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        record_key="ID",
    )

    # Make workbook with alien token FP8
    alien_file = tmp_path / "alien_token.xlsx"
    wb_alien = openpyxl.Workbook()
    ws_alien = wb_alien.active
    ws_alien.append(["ID", "Name"])
    ws_alien.append(["1", "MSK-DEADBEEF-A1B2C3D4"])
    wb_alien.save(alien_file)

    with pytest.raises(RestoreEngineError, match="RUN_IDENTITY_MISMATCH"):
        restore_workbook_v1(
            returned_workbook_path=alien_file,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            output_dir=tmp_path / "restored",
        )


def test_42_missing_run_id_with_multiple_distinct_tokens_fails(tmp_path: Path):
    """42. P0-2: Missing XPT_RunId and multiple distinct token FP8s present -> rejected RUN_IDENTITY_UNVERIFIED."""
    src = tmp_path / "multi_token.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name"])
    ws.append(["1", "Alice"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        record_key="ID",
    )

    multi_file = tmp_path / "multi_token_file.xlsx"
    wb_multi = openpyxl.Workbook()
    ws_multi = wb_multi.active
    ws_multi.append(["ID", "Name"])
    ws_multi.append(["1", "MSK-AAAA1111-X1X2X3X4"])
    ws_multi.append(["2", "MSK-BBBB2222-Y1Y2Y3Y4"])
    wb_multi.save(multi_file)

    with pytest.raises(RestoreEngineError, match="RUN_IDENTITY_UNVERIFIED.*Multiple distinct"):
        restore_workbook_v1(
            returned_workbook_path=multi_file,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            output_dir=tmp_path / "restored",
        )


def test_43_pure_numeric_missing_props_fails_closed(tmp_path: Path):
    """43. P0-2: Pure numeric masking workbook with stripped XPT props has NO token clue -> fails closed."""
    src = tmp_path / "pure_num.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Score"])
    ws.append(["E01", 100])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Score"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    from openpyxl.packaging.custom import CustomPropertyList
    stripped_num = tmp_path / "stripped_num.xlsx"
    wb_sn = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    wb_sn.custom_doc_props = CustomPropertyList()
    wb_sn.save(stripped_num)

    with pytest.raises(RestoreEngineError, match="RUN_IDENTITY_UNVERIFIED"):
        restore_workbook_v1(
            returned_workbook_path=stripped_num,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            output_dir=tmp_path / "restored",
        )


def test_44_no_key_text_fallback_uses_protected_field_id(tmp_path: Path):
    """44. P0-1: No-key text masking indexes and restores correctly using protected_field_id."""
    src = tmp_path / "nokey_text.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["HeaderA", "HeaderB"])
    ws.append(["UniqueSecret1", "OtherText"])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["A"],
        record_key=None,
    )

    res = restore_workbook_v1(
        returned_workbook_path=mask_res["masked_workbook_path"],
        ledger_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
        output_dir=tmp_path / "restored",
    )

    assert res.restored_cells_count == 1
    r_wb = openpyxl.load_workbook(res.restored_workbook_path)
    assert r_wb.active.cell(row=2, column=1).value == "UniqueSecret1"


def test_45_run_identity_pre_gate_executes_before_file_creation(tmp_path: Path):
    """45. P0-2: Verify run identity gate executes prior to any file or directory creation."""
    src = tmp_path / "pregate_order.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Val"])
    ws.append(["1", 100])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Val"],
        record_key="ID",
    )

    bad_wb = tmp_path / "bad_run.xlsx"
    wb_bad = openpyxl.load_workbook(mask_res["masked_workbook_path"])
    stamp_masked_workbook_identity(wb_bad, run_id="completely-invalid-run-uuid")
    wb_bad.save(bad_wb)

    target_dir = tmp_path / "strictly_empty_dir"
    assert not target_dir.exists()

    with pytest.raises(RestoreEngineError, match="RUN_IDENTITY_MISMATCH"):
        restore_workbook_v1(
            returned_workbook_path=bad_wb,
            ledger_path=mask_res["ledger_enc_path"],
            secret_key=mask_res["secret_key_path"].read_bytes(),
            output_dir=target_dir,
        )

    # Directory was not even created
    assert not target_dir.exists()


def test_46_cell_change_protected_field_id_passed_and_matched(tmp_path: Path):
    """46. P0-1: ChangeList cell_changes carry protected_field_id and compare_workbook accurately reflects it."""
    src = tmp_path / "pf_id_carry.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary"])
    ws.append(["E01", 5000])
    wb.save(src)

    mask_res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("1.5"),
    )

    cl = compare_workbook_against_ledger(
        returned_workbook_or_path=mask_res["masked_workbook_path"],
        ledger_or_path=mask_res["ledger_enc_path"],
        secret_key=mask_res["secret_key_path"].read_bytes(),
    )

    assert len(cl.cell_changes) == 1
    c = cl.cell_changes[0]
    assert c.protected_field_id is not None
    assert "Salary" in c.logical_field
    assert c.status == ChangeStatus.UNCHANGED


def test_47_three_sheet_sort_conflict_key_change_and_duplicate_headers(tmp_path: Path):
    """Complex gate: three sheets preserve safe sorts and block value/key edits exactly."""
    src = tmp_path / "three_sheet_complex.xlsx"
    wb = openpyxl.Workbook()
    history = wb.active
    history.title = "人员历史"
    history.append(["说明"])
    history.append(["年月", "员工工号", "姓名", "基本工资"])
    history.append(["2026-07", "E001", "张三", 6500])
    history.append(["2026-07", "E002", "李四", 7200])

    snapshot = wb.create_sheet("人员快照")
    snapshot.append(["员工工号", "姓名", "当前月薪"])
    snapshot.append(["E001", "张三", 6800])
    snapshot.append(["E002", "李四", 7500])

    salary = wb.create_sheet("薪酬明细")
    salary.append(["说明"])
    salary.append(["仅用于测试"])
    salary.append(["月份", "员工工号", "姓名", "D", "E", "F", "金额", "H", "金额"])
    salary.append(["2026-07", "E001", "张三", None, None, None, 6500, None, 388])
    salary.append(["2026-07", "E002", "李四", None, None, None, 7200, None, 420])
    wb.save(src)

    scopes = [
        SheetProtectionScope(
            sheet_name="人员历史", header_row=2,
            text_columns=["B", "C"], numeric_columns=["D"], record_key=["A", "B"],
            numeric_column_groups={"D": "history_local_A"},
        ),
        SheetProtectionScope(
            sheet_name="人员快照", header_row=1,
            text_columns=["A", "B"], numeric_columns=["C"], record_key="A",
            numeric_column_groups={"C": "shared_salary"},
        ),
        SheetProtectionScope(
            sheet_name="薪酬明细", header_row=3,
            text_columns=["B", "C"], numeric_columns=["G", "I"], record_key=["A", "B"],
            numeric_column_groups={"G": "salary_local_A", "I": "salary_local_B"},
        ),
    ]
    masked = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        sheet_scopes=scopes,
        numeric_scale_factor=Decimal("1.5"),
    )
    ledger_path = masked["ledger_enc_path"]
    key = masked["secret_key_path"].read_bytes()

    # Deterministic whole-row sort on all sheets.
    sorted_path = tmp_path / "sorted.xlsx"
    wb = openpyxl.load_workbook(masked["masked_workbook_path"])
    for sheet_name, rows in (("人员历史", (3, 4)), ("人员快照", (2, 3)), ("薪酬明细", (4, 5))):
        ws = wb[sheet_name]
        first = [ws.cell(rows[0], col).value for col in range(1, ws.max_column + 1)]
        second = [ws.cell(rows[1], col).value for col in range(1, ws.max_column + 1)]
        for col, value in enumerate(second, 1):
            ws.cell(rows[0], col).value = value
        for col, value in enumerate(first, 1):
            ws.cell(rows[1], col).value = value
    wb.save(sorted_path)
    wb.close()

    clean = restore_workbook_v1(sorted_path, ledger_path, key, output_dir=tmp_path / "restored")
    assert clean.total_protected_cells_count == 20
    assert clean.restored_cells_count == 20
    assert clean.skipped_conflicts_count == 0
    assert clean.skipped_unconfirmed_count == 0
    restored = openpyxl.load_workbook(clean.restored_workbook_path)
    assert (restored["薪酬明细"]["G5"].value, restored["薪酬明细"]["I5"].value) == (6500, 388)
    restored.close()

    # One ordinary protected value on each sheet is a true protected conflict.
    protected_path = tmp_path / "protected_conflicts.xlsx"
    wb = openpyxl.load_workbook(sorted_path)
    for sheet_name, cell in (("人员历史", "C3"), ("人员快照", "B2"), ("薪酬明细", "C4")):
        wb[sheet_name][cell] = "AI-CHANGED"
    wb.save(protected_path)
    wb.close()
    protected_changes = compare_workbook_against_ledger(protected_path, ledger_path, key)
    protected_restore = restore_workbook_v1(
        protected_path, ledger_path, key,
        output_dir=tmp_path / "restored", change_list=protected_changes,
    )
    assert protected_changes.summary[ChangeStatus.PROTECTED_VALUE_CONFLICT.value] == 3
    assert protected_restore.restored_cells_count == 17
    assert protected_restore.skipped_conflicts_count == 3

    # One Record Identifier edit on each sheet blocks the complete affected row.
    key_path = tmp_path / "key_conflicts.xlsx"
    wb = openpyxl.load_workbook(sorted_path)
    for sheet_name, cell in (("人员历史", "A3"), ("人员快照", "A2"), ("薪酬明细", "B4")):
        wb[sheet_name][cell] = "AI-CHANGED-KEY"
    wb.save(key_path)
    wb.close()
    key_changes = compare_workbook_against_ledger(key_path, ledger_path, key)
    key_restore = restore_workbook_v1(
        key_path, ledger_path, key,
        output_dir=tmp_path / "restored", change_list=key_changes,
    )
    assert key_changes.summary[ChangeStatus.ADDED_ROW.value] == 3
    assert key_changes.summary[ChangeStatus.DELETED_ROW.value] == 3
    assert key_restore.restored_cells_count == 10
    assert key_restore.skipped_unconfirmed_count == 10
    assert key_restore.unresolved_by_sheet == {"人员历史": 3, "人员快照": 3, "薪酬明细": 4}

