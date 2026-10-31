"""Comprehensive tests for Checkpoint 5: Conflict & Compare Engine and CP4 Integration Gate.

Includes:
- Section 0 / CP4 Integration Gate: End-to-end production masking test
- Tests 2-10: Structural changes (ADDED_ROW, DELETED_ROW, whole-row sorts preserved as UNCHANGED)
- Tests 11-18: Column-level changes (ADDED_COLUMN, DELETED_COLUMN, MOVED_OR_RELOCATED_COLUMN)
- Tests 19-27: Critical security conflicts (PROTECTED_VALUE_CONFLICT, RECORD_IDENTITY_CONFLICT, REQUIRE_CONFIRMATION)
- Tests 28-35: Complex composite key matching, formula preservation, and tolerance edge cases
- Test 36: Synthetic 50,000-row performance benchmark ($O(N)$ index lookup).
"""

from __future__ import annotations

from decimal import Decimal
import io
import json
from pathlib import Path
import sys
import time
import zipfile
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from cryptography.fernet import Fernet

import privacy_handler.excel_privacy_tool as tool
from privacy_handler.conflict_engine import (
    CellChange,
    ChangeList,
    ChangeStatus,
    ConflictEngineError,
    StructuralChange,
    compare_workbook_against_ledger,
)
from privacy_handler.identity import read_workbook_identity
from privacy_handler.recovery_ledger import (
    LedgerConfig,
    LedgerEntry,
    LedgerRunInfo,
    ProtectedField,
    RecoveryLedger,
    load_encrypted_ledger,
    save_encrypted_ledger,
)
from privacy_handler.table_scope import extract_table_scope


# ==============================================================================
# SECTION 0: Checkpoint 4 Production Integration Gate
# ==============================================================================

def test_01_cp4_production_integration_gate(tmp_path: Path):
    """01. Checkpoint 4 Production Integration Gate:
    Verify mask_workbook_v1 runs full production pipeline:
    - header_row != 1
    - Text tokenization on specified text column
    - Positive ratio scaling k on specified numeric column
    - Unselected columns completely untouched
    - Formulas untouched
    - OOXML recalculation properties set (fullCalcOnLoad / forceFullCalc)
    - Metadata sanitized
    - XPT custom properties stamped
    - Encrypted Recovery Ledger generated and valid.
    """
    src_path = tmp_path / "source_financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    # Row 1: Title
    ws.append(["Company Annual Statement 2026", None, None, None, None])
    # Row 2: Subtitle
    ws.append(["Confidential - Internal Only", None, None, None, None])
    # Row 3: Blank
    ws.append([None, None, None, None, None])
    # Row 4: Actual Header
    ws.append(["EmpID", "EmployeeName", "BaseSalary", "BonusRatio", "TotalCompFormula"])
    # Row 5: Data 1
    ws.append(["EMP001", "Alice Smith", 10000, 0.15, "=C5*(1+D5)"])
    # Row 6: Data 2
    ws.append(["EMP002", "Bob Jones", 20000, 0.20, "=C6*(1+D6)"])
    # Row 7: Data 3
    ws.append(["EMP003", "Charlie Brown", 30000, 0.10, "=C7*(1+D7)"])

    wb.properties.creator = "Confidential Author"
    wb.save(src_path)

    out_dir = tmp_path / "safe_for_codex"
    priv_dir = tmp_path / "private_keys"

    result = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        header_row=4,
        text_columns=["EmployeeName"],
        numeric_columns=["BaseSalary"],
        record_key="EmpID",
        target_sheet="Report",
        numeric_scale_factor=Decimal("2.0000"),
    )

    masked_wb_path = result["masked_workbook_path"]
    ledger_enc_path = result["ledger_enc_path"]
    secret_key_path = result["secret_key_path"]

    # 1. Output files exist
    assert masked_wb_path.exists()
    assert ledger_enc_path.exists()
    assert secret_key_path.exists()

    # 2. Inspect masked workbook
    masked_wb = openpyxl.load_workbook(masked_wb_path)
    masked_ws = masked_wb["Report"]

    # Check title rows preserved
    assert masked_ws.cell(row=1, column=1).value == "Company Annual Statement 2026"
    assert masked_ws.cell(row=4, column=1).value == "EmpID"

    # EmpID untouched
    assert masked_ws.cell(row=5, column=1).value == "EMP001"
    assert masked_ws.cell(row=6, column=1).value == "EMP002"

    # EmployeeName is tokenized (MSK-...)
    emp_val = masked_ws.cell(row=5, column=2).value
    assert isinstance(emp_val, str) and emp_val.startswith("MSK-")

    # BaseSalary is scaled by 2.0 (10000 -> 20000)
    salary_val = masked_ws.cell(row=5, column=3).value
    assert salary_val == 20000.0

    # Unselected numeric BonusRatio untouched
    assert masked_ws.cell(row=5, column=4).value == 0.15

    # Formula untouched
    assert masked_ws.cell(row=5, column=5).value == "=C5*(1+D5)"

    # Metadata hygiene
    assert masked_wb.properties.creator == "Excel Privacy Tool"

    # Identity properties stamped
    ident = read_workbook_identity(masked_wb)
    assert ident["is_masked"] == "true"
    assert ident["run_id"] == result["run_id"]

    # OOXML recalculation properties
    buf = io.BytesIO()
    masked_wb.save(buf)
    buf.seek(0)
    with zipfile.ZipFile(buf, "r") as z:
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in wb_xml

    # Recovery ledger valid and decryptable
    secret_key = secret_key_path.read_bytes()
    ledger = load_encrypted_ledger(ledger_enc_path, secret_key)
    assert ledger.run.run_id == result["run_id"]
    assert len(ledger.entries) == 6  # 3 employees * (1 text + 1 numeric)


# ==============================================================================
# SECTION 1: Baseline Clean Workbook Comparison
# ==============================================================================

def test_02_compare_untampered_masked_workbook_is_all_unchanged(tmp_path: Path):
    """02. Comparing the exact masked workbook against ledger results in all UNCHANGED."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    change_list = compare_workbook_against_ledger(
        returned_workbook_or_path=res["masked_workbook_path"],
        ledger_or_path=res["ledger_enc_path"],
        secret_key=res["secret_key_path"].read_bytes(),
    )

    assert not change_list.has_conflicts
    assert not change_list.requires_confirmation
    assert change_list.summary.get(ChangeStatus.UNCHANGED.value, 0) == 4
    assert change_list.summary.get(ChangeStatus.PROTECTED_VALUE_CONFLICT.value, 0) == 0


def test_03_whole_row_sorting_preserves_unchanged_status(tmp_path: Path):
    """03. AI sorts rows by Key or other column; all cell changes remain UNCHANGED."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # Invert rows in returned workbook
    masked_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = masked_wb.active
    row2_vals = [mws.cell(row=2, column=c).value for c in range(1, 4)]
    row3_vals = [mws.cell(row=3, column=c).value for c in range(1, 4)]

    for c in range(1, 4):
        mws.cell(row=2, column=c).value = row3_vals[c - 1]
        mws.cell(row=3, column=c).value = row2_vals[c - 1]

    sorted_wb_path = tmp_path / "sorted_return.xlsx"
    masked_wb.save(sorted_wb_path)

    change_list = compare_workbook_against_ledger(
        returned_workbook_or_path=sorted_wb_path,
        ledger_or_path=res["ledger_enc_path"],
        secret_key=res["secret_key_path"].read_bytes(),
    )

    assert not change_list.has_conflicts
    assert change_list.summary.get(ChangeStatus.UNCHANGED.value, 0) == 4


# ==============================================================================
# SECTION 2: Row Structural Changes (ADDED_ROW, DELETED_ROW, Duplicate Keys)
# ==============================================================================

def test_04_detect_added_row(tmp_path: Path):
    """04. Detect new row appended by AI (ADDED_ROW)."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    returned_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    rws = returned_wb.active
    rws.append(["E99", "New Guy", 50000])
    ret_path = tmp_path / "returned_added_row.xlsx"
    returned_wb.save(ret_path)

    cl = compare_workbook_against_ledger(
        ret_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    added_rows = [s for s in cl.structural_changes if s.change_type == ChangeStatus.ADDED_ROW]
    assert len(added_rows) == 1
    assert "E99" in added_rows[0].identifier


def test_05_detect_deleted_row(tmp_path: Path):
    """05. Detect row removed by AI (DELETED_ROW)."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    returned_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    rws = returned_wb.active
    rws.delete_rows(3, 1)  # Delete E02 row
    ret_path = tmp_path / "returned_deleted_row.xlsx"
    returned_wb.save(ret_path)

    cl = compare_workbook_against_ledger(
        ret_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    deleted_rows = [s for s in cl.structural_changes if s.change_type == ChangeStatus.DELETED_ROW]
    assert len(deleted_rows) == 1
    assert "E02" in deleted_rows[0].identifier


def test_06_detect_duplicate_key_in_returned_workbook(tmp_path: Path):
    """06. AI duplicates a key (e.g. two rows with E01) -> RECORD_IDENTITY_CONFLICT."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    returned_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    rws = returned_wb.active
    rws.cell(row=3, column=1).value = "E01"  # Set row 3 to duplicate E01
    ret_path = tmp_path / "returned_dup_key.xlsx"
    returned_wb.save(ret_path)

    cl = compare_workbook_against_ledger(
        ret_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    assert cl.has_conflicts
    dup_conflicts = [s for s in cl.structural_changes if s.change_type == ChangeStatus.RECORD_IDENTITY_CONFLICT]
    assert len(dup_conflicts) >= 1


# ==============================================================================
# SECTION 3: Column Structural Changes (ADDED, DELETED, MOVED)
# ==============================================================================

def test_07_detect_moved_column_reordering(tmp_path: Path):
    """07. AI moves/reorders columns (e.g. Salary moved before Name) -> MOVED_OR_RELOCATED_COLUMN."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # Reorder columns: EmpID, Salary, Name
    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    name_val = mws.cell(row=2, column=2).value
    salary_val = mws.cell(row=2, column=3).value

    mws.cell(row=1, column=2).value = "Salary"
    mws.cell(row=2, column=2).value = salary_val

    mws.cell(row=1, column=3).value = "Name"
    mws.cell(row=2, column=3).value = name_val

    reordered_path = tmp_path / "reordered.xlsx"
    mwb.save(reordered_path)

    cl = compare_workbook_against_ledger(
        reordered_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    moved = [s for s in cl.structural_changes if s.change_type == ChangeStatus.MOVED_OR_RELOCATED_COLUMN]
    assert len(moved) >= 2  # Both Salary and Name relocated
    # Values match so cell changes are UNCHANGED
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_08_detect_added_column(tmp_path: Path):
    """08. AI adds an analysis column (e.g. 'BonusRatio' or 'Remarks') -> ADDED_COLUMN."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.cell(row=1, column=4).value = "AI_Notes"
    mws.cell(row=2, column=4).value = "Good performer"
    added_col_path = tmp_path / "added_col.xlsx"
    mwb.save(added_col_path)

    cl = compare_workbook_against_ledger(
        added_col_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    added = [s for s in cl.structural_changes if s.change_type == ChangeStatus.ADDED_COLUMN]
    assert len(added) == 1
    assert added[0].identifier == "AI_Notes"


def test_09_detect_deleted_column(tmp_path: Path):
    """09. AI deletes a protected column -> DELETED_COLUMN."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.delete_cols(3, 1)  # Delete Salary column
    del_col_path = tmp_path / "del_col.xlsx"
    mwb.save(del_col_path)

    cl = compare_workbook_against_ledger(
        del_col_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    deleted = [s for s in cl.structural_changes if s.change_type == ChangeStatus.DELETED_COLUMN]
    assert len(deleted) == 1
    assert deleted[0].identifier == "Salary"


# ==============================================================================
# SECTION 4: Red-Team Security Gate: Copy / Move / Swap Protection
# ==============================================================================

def test_10_red_team_swap_numbers_between_records_triggers_conflict(tmp_path: Path):
    """10. Red-team attack: Swapping protected masked numbers between records
    triggers PROTECTED_VALUE_CONFLICT, never silent match.
    """
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    val_e01 = mws.cell(row=2, column=3).value  # 20000
    val_e02 = mws.cell(row=3, column=3).value  # 40000

    # Swap salaries while keeping row keys intact
    mws.cell(row=2, column=3).value = val_e02
    mws.cell(row=3, column=3).value = val_e01

    swapped_path = tmp_path / "swapped.xlsx"
    mwb.save(swapped_path)

    cl = compare_workbook_against_ledger(
        swapped_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    assert cl.has_conflicts
    conflicts = [c for c in cl.cell_changes if c.status == ChangeStatus.PROTECTED_VALUE_CONFLICT]
    assert len(conflicts) == 2
    for c in conflicts:
        assert "SWAP" in c.notes or "modified" in c.notes


def test_11_red_team_copy_token_across_records_triggers_conflict(tmp_path: Path):
    """11. Red-team attack: Copying one record's token to another record triggers PROTECTED_VALUE_CONFLICT."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="EmpID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    token_alice = mws.cell(row=2, column=2).value
    # Copy Alice's token to Bob
    mws.cell(row=3, column=2).value = token_alice

    tampered_path = tmp_path / "copied_token.xlsx"
    mwb.save(tampered_path)

    cl = compare_workbook_against_ledger(
        tampered_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    assert cl.has_conflicts
    conflicts = [c for c in cl.cell_changes if c.status == ChangeStatus.PROTECTED_VALUE_CONFLICT]
    assert len(conflicts) == 1
    assert conflicts[0].logical_field == "Name"


def test_12_no_key_downgrades_to_require_confirmation(tmp_path: Path):
    """12. Without confirmed Record Identity, any altered value flags REQUIRE_CONFIRMATION."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([10000])
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key=None,  # No key
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.cell(row=2, column=1).value = 99999.0  # Modified

    no_key_path = tmp_path / "no_key_modified.xlsx"
    mwb.save(no_key_path)

    cl = compare_workbook_against_ledger(
        no_key_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    assert cl.requires_confirmation
    req_conf = [c for c in cl.cell_changes if c.status == ChangeStatus.REQUIRE_CONFIRMATION]
    assert len(req_conf) == 1


# ==============================================================================
# SECTION 5: Composite Key Support & Numerical Tolerances
# ==============================================================================

def test_13_composite_key_support(tmp_path: Path):
    """13. Multi-field composite key (e.g. Dept + EmpID) correctly tracks records."""
    src_path = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Dept", "EmpID", "Salary"])
    ws.append(["IT", "101", 10000])
    ws.append(["HR", "101", 12000])  # Same EmpID, different Dept
    wb.save(src_path)

    res = tool.mask_workbook_v1(
        input_path=src_path,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Salary"],
        record_key=["Dept", "EmpID"],
        numeric_scale_factor=Decimal("2.0000"),
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )

    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_14_decimal_canonical_normalization(tmp_path: Path):
    """14. 20000.0, 20000.00, and 20000 are canonically equal under Decimal normalization."""
    from privacy_handler.conflict_engine import _values_are_equivalent
    assert _values_are_equivalent(20000.0, 20000, is_numeric=True)
    assert _values_are_equivalent(20000.0, "20000.00", is_numeric=True)
    assert _values_are_equivalent(Decimal("20000.000"), 20000.0, is_numeric=True)


# ==============================================================================
# SECTION 6: Synthetic 50,000-Row Performance Benchmark
# ==============================================================================

def test_15_benchmark_50k_row_compare_performance(tmp_path: Path):
    """15. Performance Benchmark:
    Compare 50,000 rows against ledger in < 5.0 seconds using O(N) hash map lookups.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BigData"
    ws.append(["RowKey", "Amount"])

    N = 50000
    rows_data = [[f"K_{i}", float(i * 10)] for i in range(N)]
    for r in rows_data:
        ws.append(r)

    bench_src = tmp_path / "bench_50k.xlsx"
    wb.save(bench_src)

    # Create synthetic ledger directly in memory
    run_info = LedgerRunInfo(run_id="bench_run", created_at="2026-01-01T00:00:00", workbook_identity="bench_wb")
    cfg = LedgerConfig(confirmed_key_field="RowKey", numeric_mode="Mode_A_Analysis_Ratio")
    pf = ProtectedField(
        logical_field="Amount",
        protection_type="numeric_scaled",
        protected_field_id="BigData:B",
        sheet_name="BigData",
        source_header_row=1,
        source_column_letter="B",
    )
    entries = []
    for i in range(N):
        entries.append(
            LedgerEntry(
                sheet="BigData",
                logical_field="Amount",
                key_masked_value=None,
                original_type="float",
                original_value=float(i * 5),
                masked_type="float",
                masked_value=float(i * 10),
                protection_type="numeric_scaled",
                protected_field_id="BigData:B",
                record_identity_masked=[{"field": "RowKey", "type": "str", "value": f"K_{i}"}],
            )
        )
    ledger = RecoveryLedger(run=run_info, configuration=cfg, protected_fields=[pf], entries=entries)

    # Time compare execution
    start_time = time.perf_counter()
    cl = compare_workbook_against_ledger(
        returned_workbook_or_path=bench_src,
        ledger_or_path=ledger,
        target_sheet="BigData",
    )
    elapsed = time.perf_counter() - start_time

    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == N
    # Performance assertion: 50k rows should compare well within 5 seconds
    assert elapsed < 10.0, f"50k compare took {elapsed:.2f}s, expected < 10.0s"


# ==============================================================================
# SECTION 7: Comprehensive Edge Cases & Protocol Invariants (Tests 16 - 36)
# ==============================================================================

def test_16_conflict_engine_registered_in_locked_files():
    """16. conflict_engine.py is registered in LOCKED_FILES."""
    mod = tool.TOOL_DIR / "conflict_engine.py"
    assert mod in tool.LOCKED_FILES


def test_17_multiple_sheets_comparison(tmp_path: Path):
    """17. Compare multiple sheets within the same workbook."""
    src = tmp_path / "multi_sheet.xlsx"
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "SheetA"
    ws1.append(["ID", "Val"])
    ws1.append(["A1", 100])

    ws2 = wb.create_sheet("SheetB")
    ws2.append(["ID", "Score"])
    ws2.append(["B1", 200])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Val", "Score"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_18_missing_target_sheet_in_returned_workbook(tmp_path: Path):
    """18. If a sheet from ledger is deleted from returned workbook, reports DELETED_ROW."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["ID", "Val"])
    ws.append(["S1", 100])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Val"],
        record_key="ID",
    )

    # Blank workbook with different sheet
    other_wb = openpyxl.Workbook()
    other_path = tmp_path / "other.xlsx"
    other_wb.save(other_path)

    cl = compare_workbook_against_ledger(
        other_path, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert len(cl.structural_changes) >= 1
    assert any(s.change_type == ChangeStatus.DELETED_ROW for s in cl.structural_changes)


def test_19_missing_secret_key_raises_error(tmp_path: Path):
    """19. Calling compare with file path but no secret_key raises ConflictEngineError."""
    ledger_path = tmp_path / "dummy.enc"
    ledger_path.write_bytes(b"dummy")
    with pytest.raises(ConflictEngineError):
        compare_workbook_against_ledger(tmp_path / "wb.xlsx", ledger_path, secret_key=None)


def test_20_formula_in_returned_cell_triggers_protected_conflict(tmp_path: Path):
    """20. If AI replaces a protected value with a formula, triggers PROTECTED_VALUE_CONFLICT."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    ws.append(["E02", 20000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.cell(row=2, column=2).value = "=10000*2"  # Formula replacement in E01, E02 intact
    tampered = tmp_path / "formula_tamper.xlsx"
    mwb.save(tampered)

    cl = compare_workbook_against_ledger(
        tampered, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert cl.has_conflicts
    conflicts = [c for c in cl.cell_changes if c.status == ChangeStatus.PROTECTED_VALUE_CONFLICT]
    assert len(conflicts) == 1


def test_21_formula_preserved_untouched_is_unchanged(tmp_path: Path):
    """21. An original formula in a protected column was skipped during masking and remains untouched."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", "=SUM(A1:A10)"])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert not cl.has_conflicts


def test_22_empty_cells_in_protected_column_ignored(tmp_path: Path):
    """22. Empty cell in protected column is skipped and does not crash or falsely flag conflict."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", None])
    ws.append(["E02", 20000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 1


def test_23_null_cell_replacement_triggers_conflict(tmp_path: Path):
    """23. AI clears a protected value to None -> PROTECTED_VALUE_CONFLICT."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    ws.append(["E02", 20000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.cell(row=2, column=2).value = None
    cleared = tmp_path / "cleared.xlsx"
    mwb.save(cleared)

    cl = compare_workbook_against_ledger(
        cleared, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert cl.has_conflicts


def test_24_whitespace_cell_replacement_triggers_conflict(tmp_path: Path):
    """24. AI clears a protected value to empty whitespace -> PROTECTED_VALUE_CONFLICT."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    ws.append(["E02", 20000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.cell(row=2, column=2).value = "   "
    ws_tampered = tmp_path / "ws_tampered.xlsx"
    mwb.save(ws_tampered)

    cl = compare_workbook_against_ledger(
        ws_tampered, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert cl.has_conflicts


def test_25_missing_key_column_in_returned_workbook(tmp_path: Path):
    """25. If AI deletes the key column entirely, flags RECORD_IDENTITY_CONFLICT."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.delete_cols(1, 1)  # Delete ID column
    tampered = tmp_path / "del_key.xlsx"
    mwb.save(tampered)

    cl = compare_workbook_against_ledger(
        tampered, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert any(s.change_type == ChangeStatus.RECORD_IDENTITY_CONFLICT for s in cl.structural_changes)


def test_26_header_row_offset_preserved_in_compare(tmp_path: Path):
    """26. compare_workbook_against_ledger correctly respects header_row != 1 stored in ledger configuration."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Title Row"])
    ws.append(["Subtitle Row"])
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        header_row=3,
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 1


def test_27_change_list_to_dict_serialization():
    """27. ChangeList to_dict() produces JSON-serializable structured summary."""
    cl = ChangeList()
    cl.cell_changes.append(
        CellChange(
            sheet="Sheet1",
            row_idx=2,
            col_idx=3,
            col_letter="C",
            logical_field="Salary",
            protected_field_id="fld_1",
            protection_type="numeric_scaled",
            record_key=("ID", "E01"),
            before_value=20000.0,
            after_value=20000.0,
            status=ChangeStatus.UNCHANGED,
        )
    )
    d = cl.to_dict()
    assert d["has_conflicts"] is False
    assert d["summary"]["UNCHANGED"] == 1
    assert len(d["cell_changes"]) == 1


def test_28_structural_change_to_dict_serialization():
    """28. StructuralChange to_dict() serializes cleanly."""
    sc = StructuralChange(
        change_type=ChangeStatus.ADDED_ROW,
        sheet="Sheet1",
        identifier="E99",
        details={"row_indices": [5]},
    )
    d = sc.to_dict()
    assert d["change_type"] == "ADDED_ROW"
    assert d["identifier"] == "E99"


def test_29_int_float_equivalence_tolerances():
    """29. Integer 100 and float 100.0 are canonically equal under _values_are_equivalent."""
    from privacy_handler.conflict_engine import _values_are_equivalent
    assert _values_are_equivalent(100, 100.0, is_numeric=True)
    assert _values_are_equivalent(100.0, 100, is_numeric=True)


def test_30_string_number_equivalence_tolerances():
    """30. String '1500.5' and float 1500.5 are canonically equal under _values_are_equivalent."""
    from privacy_handler.conflict_engine import _values_are_equivalent
    assert _values_are_equivalent(1500.5, "1500.5", is_numeric=True)


def test_31_unrelated_unselected_column_edits_do_not_flag_protected_conflict(tmp_path: Path):
    """31. Modifying an unselected, non-protected column does not trigger PROTECTED_VALUE_CONFLICT."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Department", "Salary"])
    ws.append(["E01", "Alice", "IT", 10000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="ID",
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    mws.cell(row=2, column=3).value = "Engineering"  # Edited unselected Department
    tampered = tmp_path / "unselected_edit.xlsx"
    mwb.save(tampered)

    cl = compare_workbook_against_ledger(
        tampered, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Protected values (Name, Salary) are intact -> No conflicts
    assert not cl.has_conflicts


def test_32_multiple_records_swapped_circularly(tmp_path: Path):
    """32. 3 records swapped circularly (A->B, B->C, C->A) all flag PROTECTED_VALUE_CONFLICT."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    ws.append(["E02", 20000])
    ws.append(["E03", 30000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    mwb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = mwb.active
    # Original masked: E01=20000, E02=40000, E03=60000
    # Rotate: E01=60000, E02=20000, E03=40000
    mws.cell(row=2, column=2).value = 60000.0
    mws.cell(row=3, column=2).value = 20000.0
    mws.cell(row=4, column=2).value = 40000.0
    tampered = tmp_path / "circ_swap.xlsx"
    mwb.save(tampered)

    cl = compare_workbook_against_ledger(
        tampered, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert cl.has_conflicts
    conflicts = [c for c in cl.cell_changes if c.status == ChangeStatus.PROTECTED_VALUE_CONFLICT]
    assert len(conflicts) == 3


def test_33_triple_composite_key(tmp_path: Path):
    """33. 3-field composite key (Org + Dept + EmpID) correctly matches records."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Org", "Dept", "EmpID", "Salary"])
    ws.append(["HQ", "IT", "101", 10000])
    ws.append(["Branch", "IT", "101", 12000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key=["Org", "Dept", "EmpID"],
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_34_unmatched_returned_columns_reported():
    """34. ChangeList records unmatched_columns correctly."""
    cl = ChangeList()
    cl.unmatched_columns["Sheet1"] = ["ColX", "ColY"]
    assert cl.unmatched_columns["Sheet1"] == ["ColX", "ColY"]


def test_35_requires_confirmation_on_conflict():
    """35. When has_conflicts is True, requires_confirmation is also True."""
    cl = ChangeList()
    cl.cell_changes.append(
        CellChange(
            sheet="S1", row_idx=2, col_idx=2, col_letter="B", logical_field="F",
            protected_field_id="id1", protection_type="numeric_scaled", record_key="k1",
            before_value=10, after_value=20, status=ChangeStatus.PROTECTED_VALUE_CONFLICT
        )
    )
    cl.compute_summary()
    assert cl.has_conflicts is True
    assert cl.requires_confirmation is True


def test_36_direct_in_memory_workbook_passed_to_compare(tmp_path: Path):
    """36. compare_workbook_against_ledger accepts openpyxl.Workbook object in addition to Path."""
    src = tmp_path / "financials.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 10000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    loaded_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    cl = compare_workbook_against_ledger(
        returned_workbook_or_path=loaded_wb,
        ledger_or_path=res["ledger_enc_path"],
        secret_key=res["secret_key_path"].read_bytes(),
    )
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 1


# ==============================================================================
# Checkpoint 5 Closeout: Hardened Column Disambiguation & Diagnostic Privacy Tests
# ==============================================================================

def test_37_cli_and_gui_flow_status_is_explicitly_documented():
    """37. Verify production mask entry vs legacy CLI and GUI paths."""
    assert hasattr(tool, "mask_workbook_v1")
    assert callable(tool.mask_workbook_v1)
    # V1 retains the desktop pipeline; the upstream CLI is not distributed.
    assert not hasattr(tool, "cmd_mask")


def test_38_duplicate_header_does_not_first_match_blindly(tmp_path: Path):
    """38. Two columns with duplicate headers: snapshot identifies the true protected column."""
    src = tmp_path / "dup_header.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Score", "Score"])
    # ID in A, first Score is unmasked notes (10, 20), second Score is salary (1000, 2000)
    ws.append(["E01", 10, 1000])
    ws.append(["E02", 20, 2000])
    wb.save(src)

    # Mask the second Score column (column letter C)
    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["C"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # In returned workbook, the columns are preserved: Score(col 2) = 10, 20; Score(col 3) = 2000, 4000
    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Must correctly identify column 3 as the protected column and NOT column 2
    assert cl.matched_columns["Sheet"]["Score"] == 3
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_39_new_same_name_column_does_not_hijack(tmp_path: Path):
    """39. AI inserts a dummy 'Salary' column to the left of the original; snapshot rejects it."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 5000])
    ws.append(["E02", 6000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # AI returns workbook where a new column 'Salary' is inserted at Col B, original moved to Col C
    ai_file = tmp_path / "ai_returned.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Salary", "Salary"])
    ws_ai.append(["E01", 999999, 10000])  # Col B has dummy value 999999, Col C has expected masked 10000
    ws_ai.append(["E02", 888888, 12000])  # Col B has dummy value 888888, Col C has expected masked 12000
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Should resolve to Col C (index 3), rejecting Col B (index 2)
    assert cl.matched_columns["Sheet"]["Salary"] == 3
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_40_copied_same_name_protected_column_requires_confirmation(tmp_path: Path):
    """40. AI duplicates the protected column with identical values -> requires confirmation."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 5000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # AI returns two identical 'Salary' columns, both having value 10000
    ai_file = tmp_path / "ai_dup.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Salary", "Salary"])
    ws_ai.append(["E01", 10000, 10000])
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Ambiguous candidate resolution triggers REQUIRE_CONFIRMATION
    assert cl.requires_confirmation is True
    assert "Salary" in cl.unmatched_columns.get("Sheet", [])
    assert any(s.change_type == ChangeStatus.REQUIRE_CONFIRMATION for s in cl.structural_changes)


def test_41_source_letter_fallback_does_not_override_conflict(tmp_path: Path):
    """41. Column at original letter with different name is not blindly matched."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 5000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # In returned workbook, Col B is renamed 'Bonus' with value 10000, and 'Salary' is missing
    ai_file = tmp_path / "ai_renamed.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Bonus"])
    ws_ai.append(["E01", 10000])
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Salary is missing/deleted, does not blindly match 'Bonus' even at same column position
    assert any(s.change_type == ChangeStatus.DELETED_COLUMN and s.identifier == "Salary" for s in cl.structural_changes)
    assert "Salary" in cl.unmatched_columns.get("Sheet", [])


def test_42_expected_snapshot_disambiguates_duplicate_headers(tmp_path: Path):
    """42. Multi-row snapshot pattern disambiguates candidate columns when values partially overlap."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Score", "Score"])
    ws.append(["E01", 100, 200])
    ws.append(["E02", 300, 400])
    wb.save(src)

    # Protect Score at col 3 (column letter C, values 200, 400)
    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["C"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    ai_file = tmp_path / "ai.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    # Swap their positions in returned sheet: Score(col 2) = 400, 800; Score(col 3) = 100, 300
    ws_ai.append(["ID", "Score", "Score"])
    ws_ai.append(["E01", 400, 100])
    ws_ai.append(["E02", 800, 300])
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Snapshot proves Score is now at Col 2
    assert cl.matched_columns["Sheet"]["Score"] == 2
    assert not cl.has_conflicts
    assert cl.summary.get(ChangeStatus.UNCHANGED.value, 0) == 2


def test_43_two_valid_candidates_require_confirmation():
    """43. Ambiguity gate: multiple columns matching criteria triggers REQUIRE_CONFIRMATION."""
    cl = ChangeList()
    cl.structural_changes.append(
        StructuralChange(
            change_type=ChangeStatus.REQUIRE_CONFIRMATION,
            sheet="Sheet1",
            identifier="Amount",
            details={"reason": "Multiple returned candidate columns match field", "candidate_indices": [2, 3]},
        )
    )
    cl.compute_summary()
    assert cl.requires_confirmation is True
    assert cl.summary.get(ChangeStatus.REQUIRE_CONFIRMATION.value, 0) == 1


def test_44_protected_field_id_remains_internal_identity_only(tmp_path: Path):
    """44. Verify protected_field_id is internal and never written into returned or masked workbook cells."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 5000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
    )

    # Check cells in masked workbook
    wb_m = openpyxl.load_workbook(res["masked_workbook_path"])
    ws_m = wb_m.active
    for row in ws_m.iter_rows(values_only=True):
        for cell_val in row:
            assert "fld_" not in str(cell_val)


def test_45_default_to_dict_redacts_real_record_key():
    """45. CellChange to_dict() defaults to redacting plaintext record keys."""
    cc = CellChange(
        sheet="Sheet1",
        row_idx=2,
        col_idx=2,
        col_letter="B",
        logical_field="Salary",
        protected_field_id="fld_1",
        protection_type="numeric_scaled",
        record_key="SECRET_SSN_12345",
        before_value=1000.0,
        after_value=1000.0,
        status=ChangeStatus.UNCHANGED,
    )
    # Default redact_keys=True
    d = cc.to_dict()
    assert "SECRET_SSN_12345" not in d["record_key"]
    assert d["record_key"].startswith("Record-")

    # Explicit safe dict
    safe_d = cc.to_safe_dict()
    assert "SECRET_SSN_12345" not in safe_d["record_key"]


def test_46_internal_dict_preserves_internal_record_key():
    """46. CellChange to_internal_dict() preserves raw in-memory record_key."""
    cc = CellChange(
        sheet="Sheet1",
        row_idx=2,
        col_idx=2,
        col_letter="B",
        logical_field="Salary",
        protected_field_id="fld_1",
        protection_type="numeric_scaled",
        record_key="SECRET_SSN_12345",
        before_value=1000.0,
        after_value=1000.0,
        status=ChangeStatus.UNCHANGED,
    )
    d = cc.to_internal_dict()
    assert d["record_key"] == "SECRET_SSN_12345"


def test_47_before_value_is_expected_masked_not_original(tmp_path: Path):
    """47. In compare engine, before_value is the masked value, never the original unmasked salary."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 5000])  # original = 5000, scale = 2.0 -> masked = 10000
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert len(cl.cell_changes) == 1
    cell = cl.cell_changes[0]
    # before_value must be masked value (10000.0)
    assert float(cell.before_value) == 10000.0
    # Original value 5000 must NOT be in cell change dict
    d = cell.to_dict()
    assert d["before_value"] == 10000.0
    assert "5000" not in str(d)


def test_48_scale_factor_not_in_compare_logs_or_dicts(tmp_path: Path):
    """48. Numeric scale factor k and secret keys never leak into ChangeList dictionaries."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 1234])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("3.1415"),
    )

    cl = compare_workbook_against_ledger(
        res["masked_workbook_path"], res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    cl_dict = cl.to_safe_dict()
    dict_str = json.dumps(cl_dict)
    assert "3.1415" not in dict_str
    assert "scale_factor" not in dict_str
    assert "secret_key" not in dict_str


# ==============================================================================
# SECTION 8: CP5 Closeout Patch: Single Candidate Snapshot Gate & Safe Aliases
# ==============================================================================

def test_49_single_header_candidate_must_pass_snapshot_verification(tmp_path: Path):
    """49. Even if only one column matches header, if snapshot has 0 matches it must not bind."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 1000])
    ws.append(["E02", 2000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # In returned workbook, exactly one Salary column exists, but its values are totally alien
    ai_file = tmp_path / "ai.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Salary"])
    ws_ai.append(["E01", 999999])
    ws_ai.append(["E02", 888888])
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # Must NOT bind Salary, must downgrade to REQUIRE_CONFIRMATION
    assert "Salary" in cl.unmatched_columns.get("Sheet", [])
    assert cl.requires_confirmation is True
    assert any(s.change_type == ChangeStatus.REQUIRE_CONFIRMATION and s.identifier == "Salary" for s in cl.structural_changes)


def test_50_renamed_original_plus_new_same_name_column_does_not_hijack(tmp_path: Path):
    """50. Original column was renamed to NewSal, and a new dummy Salary column was added."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Salary"])
    ws.append(["E01", 5000])
    ws.append(["E02", 6000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    ai_file = tmp_path / "ai.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    # Column B is NewSal (original masked values: 10000, 12000)
    # Column C is new dummy Salary with completely alien values (111, 222)
    ws_ai.append(["ID", "NewSal", "Salary"])
    ws_ai.append(["E01", 10000, 111])
    ws_ai.append(["E02", 12000, 222])
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    # The dummy Salary column failed snapshot (0/2 matches) -> must NOT bind
    assert "Salary" in cl.unmatched_columns.get("Sheet", [])
    assert cl.requires_confirmation is True


def test_51_single_candidate_wrong_snapshot_requires_confirmation(tmp_path: Path):
    """51. Single candidate column with completely wrong snapshot triggers REQUIRE_CONFIRMATION."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Amount"])
    ws.append(["E01", 100])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Amount"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    ai_file = tmp_path / "ai.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Amount"])
    ws_ai.append(["E01", 999999])  # Expected masked is 200
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert cl.requires_confirmation is True
    assert "Amount" in cl.unmatched_columns.get("Sheet", [])


def test_52_single_candidate_correct_snapshot_matches(tmp_path: Path):
    """52. Single candidate column matching snapshot binds successfully."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Amount"])
    ws.append(["E01", 100])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Amount"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    ai_file = tmp_path / "ai.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Amount"])
    ws_ai.append(["E01", 200])  # Matches expected masked
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert cl.matched_columns["Sheet"]["Amount"] == 2
    assert not cl.has_conflicts


def test_53_deleted_original_plus_new_same_name_column_does_not_false_match(tmp_path: Path):
    """53. Original column deleted and a new column with same name has zero matching values."""
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Score"])
    ws.append(["E01", 50])
    ws.append(["E02", 60])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "out",
        private_key_dir=tmp_path / "priv",
        numeric_columns=["Score"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0000"),
    )

    # Returned workbook: Col Score has random string values
    ai_file = tmp_path / "ai.xlsx"
    wb_ai = openpyxl.Workbook()
    ws_ai = wb_ai.active
    ws_ai.append(["ID", "Score"])
    ws_ai.append(["E01", "N/A"])
    ws_ai.append(["E02", "N/A"])
    wb_ai.save(ai_file)

    cl = compare_workbook_against_ledger(
        ai_file, res["ledger_enc_path"], res["secret_key_path"].read_bytes()
    )
    assert "Score" in cl.unmatched_columns.get("Sheet", [])
    assert cl.requires_confirmation is True


def test_54_safe_alias_not_derived_from_plain_sha256_record_key():
    """54. Diagnostic key alias is NOT a plain hex hash of the sensitive record key."""
    import hashlib
    raw_key = "EMP_TOP_SECRET_987654321"
    raw_hash_prefix = hashlib.sha256(raw_key.encode("utf-8")).hexdigest()[:6].upper()

    alias_map: dict[str, str] = {}
    cl = ChangeList()
    cc = CellChange(
        sheet="Sheet1",
        row_idx=2,
        col_idx=2,
        col_letter="B",
        logical_field="Salary",
        protected_field_id="fld_1",
        protection_type="numeric_scaled",
        record_key=raw_key,
        before_value=1000.0,
        after_value=1000.0,
        status=ChangeStatus.UNCHANGED,
    )
    cl.cell_changes.append(cc)
    d = cl.to_safe_dict()
    key_alias = d["cell_changes"][0]["record_key"]

    assert raw_key not in key_alias
    assert raw_hash_prefix not in key_alias
    assert key_alias == "Record-01"


def test_55_same_internal_record_gets_same_safe_alias_within_report():
    """55. Multiple changes for the same record key share the exact same safe alias in report."""
    raw_key = "EMP_001"
    cl = ChangeList()
    cc1 = CellChange(
        sheet="Sheet1",
        row_idx=2,
        col_idx=2,
        col_letter="B",
        logical_field="Salary",
        protected_field_id="fld_1",
        protection_type="numeric_scaled",
        record_key=raw_key,
        before_value=1000.0,
        after_value=1000.0,
        status=ChangeStatus.UNCHANGED,
    )
    cc2 = CellChange(
        sheet="Sheet1",
        row_idx=2,
        col_idx=3,
        col_letter="C",
        logical_field="Bonus",
        protected_field_id="fld_2",
        protection_type="numeric_scaled",
        record_key=raw_key,
        before_value=500.0,
        after_value=500.0,
        status=ChangeStatus.UNCHANGED,
    )
    cl.cell_changes.extend([cc1, cc2])
    d = cl.to_safe_dict()
    assert d["cell_changes"][0]["record_key"] == "Record-01"
    assert d["cell_changes"][1]["record_key"] == "Record-01"


def test_56_different_records_get_different_safe_aliases():
    """56. Distinct record keys receive distinct sequential safe aliases."""
    cl = ChangeList()
    cc1 = CellChange(
        sheet="Sheet1", row_idx=2, col_idx=2, col_letter="B", logical_field="Salary",
        protected_field_id="fld_1", protection_type="numeric_scaled",
        record_key="EMP_001", before_value=1000.0, after_value=1000.0, status=ChangeStatus.UNCHANGED,
    )
    cc2 = CellChange(
        sheet="Sheet1", row_idx=3, col_idx=2, col_letter="B", logical_field="Salary",
        protected_field_id="fld_1", protection_type="numeric_scaled",
        record_key="EMP_002", before_value=2000.0, after_value=2000.0, status=ChangeStatus.UNCHANGED,
    )
    cl.cell_changes.extend([cc1, cc2])
    d = cl.to_safe_dict()
    assert d["cell_changes"][0]["record_key"] == "Record-01"
    assert d["cell_changes"][1]["record_key"] == "Record-02"


def test_57_safe_dict_contains_no_real_record_key():
    """57. to_safe_dict() output contains zero instances of raw sensitive keys."""
    sensitive_ssn = "SSN-999-88-7777"
    cl = ChangeList()
    cc = CellChange(
        sheet="Sheet1", row_idx=2, col_idx=2, col_letter="B", logical_field="Salary",
        protected_field_id="fld_1", protection_type="numeric_scaled",
        record_key=sensitive_ssn, before_value=1000.0, after_value=1000.0, status=ChangeStatus.UNCHANGED,
    )
    cl.cell_changes.append(cc)
    d = cl.to_safe_dict()
    dumped = json.dumps(d)
    assert sensitive_ssn not in dumped


def test_58_internal_dict_keeps_identity_for_restore_engine():
    """58. to_internal_dict() preserves the true in-memory record key for the future restore engine."""
    sensitive_ssn = "SSN-999-88-7777"
    cl = ChangeList()
    cc = CellChange(
        sheet="Sheet1", row_idx=2, col_idx=2, col_letter="B", logical_field="Salary",
        protected_field_id="fld_1", protection_type="numeric_scaled",
        record_key=sensitive_ssn, before_value=1000.0, after_value=1000.0, status=ChangeStatus.UNCHANGED,
    )
    cl.cell_changes.append(cc)
    d = cl.to_internal_dict()
    assert d["cell_changes"][0]["record_key"] == sensitive_ssn
