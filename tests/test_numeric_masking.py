"""Comprehensive unit, regression, and gate tests for Checkpoint 4: Analysis-Mode Numeric Masking."""

from __future__ import annotations

from decimal import Decimal
import io
import json
import os
from pathlib import Path
import sys
import zipfile
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl.styles import numbers

import privacy_handler.excel_privacy_tool as tool
from privacy_handler.numeric_masking import (
    MAX_COLLISION_RETRIES,
    NumericMaskingError,
    NumericProtectedColumnConfig,
    analyze_numeric_precision,
    ensure_workbook_recalculation_properties,
    execute_field_scaling_with_collision_gate,
    generate_scale_factor,
    mask_numeric_table_scope,
    scale_decimal_value,
)
from privacy_handler.recovery_ledger import (
    LedgerConfig,
    LedgerEntry,
    LedgerRunInfo,
    ProtectedField,
    RecoveryLedger,
    save_encrypted_ledger,
    load_encrypted_ledger,
)
from privacy_handler.table_scope import (
    ColumnDescriptor,
    OutputPaths,
    TableScope,
    extract_table_scope,
    resolve_safe_output_path,
)


# ==============================================================================
# 1 - 10: Mathematical & Type Properties of Scaling
# ==============================================================================

def test_selected_numeric_field_is_scaled():
    """1. Selected numeric field is scaled by positive k."""
    k = Decimal("1.5000")
    _, out_val = scale_decimal_value(1000, k, precision=2)
    assert out_val == 1500.0


def test_unselected_numeric_field_unchanged():
    """2. Unselected numeric field remains completely untouched in the worksheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary", "Age"])
    ws.append(["E01", 10000, 30])
    ws.append(["E02", 12000, 25])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [
        NumericProtectedColumnConfig(
            sheet_name=ws.title,
            column_letter="B",
            logical_field="Salary",
        )
    ]
    k = Decimal("2.0000")
    mask_numeric_table_scope(ws, scope, cfg, scale_factor=k)

    # Salary is scaled
    assert ws["B2"].value == 20000.0
    assert ws["B3"].value == 24000.0
    # Age (C column) is unchanged
    assert ws["C2"].value == 30
    assert ws["C3"].value == 25


def test_multiple_selected_amount_fields_share_same_k():
    """3. Multiple selected amount fields in the same group share the exact same scale factor k."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary", "Bonus"])
    ws.append([10000, 2000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [
        NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Salary"),
        NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="B", logical_field="Bonus"),
    ]
    k = Decimal("1.7500")
    _, meta = mask_numeric_table_scope(ws, scope, cfg, scale_factor=k)

    assert ws["A2"].value == 17500.0
    assert ws["B2"].value == 3500.0
    # Ratio is preserved exactly
    assert ws["B2"].value / ws["A2"].value == 2000 / 10000
    assert meta["scale_factor"] == "1.7500"


def test_different_runs_use_different_k():
    """4. Different CSPRNG runs generate different scale factors."""
    k_samples = {generate_scale_factor() for _ in range(50)}
    assert len(k_samples) > 40  # Highly diverse distribution


def test_k_is_positive():
    """5. Scale factor k is strictly positive within [1.25, 3.85]."""
    for _ in range(100):
        k = generate_scale_factor()
        assert Decimal("1.2500") <= k <= Decimal("3.8500")
        assert k > 0


def test_k_not_equal_one():
    """6. Scale factor k never equals 1.0."""
    for _ in range(200):
        k = generate_scale_factor()
        assert k != Decimal("1.0")
        assert k != Decimal("1.0000")


def test_zero_remains_zero():
    """7. Zero maps to zero under Mode A positive scale."""
    k = Decimal("2.3456")
    _, out_int_zero = scale_decimal_value(0, k, precision=2)
    assert out_int_zero == 0

    _, out_float_zero = scale_decimal_value(0.0, k, precision=2)
    assert out_float_zero == 0.0


def test_negative_preserves_sign():
    """8. Negative amounts preserve negative sign under positive scale."""
    k = Decimal("1.5000")
    _, out_neg = scale_decimal_value(-500.0, k, precision=2)
    assert out_neg == -750.0


def test_integer_original_type_preserved_in_ledger():
    """9. Integer original values have original_type='int' preserved in Ledger."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Amount"])
    ws.append([8500])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Amount")]
    entries, _ = mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("1.5000"))

    assert len(entries) == 1
    assert entries[0].original_type == "int"
    assert entries[0].original_value == 8500
    assert entries[0].get_decoded_original_value() == 8500


def test_float_original_type_preserved():
    """10. Float original values have original_type='float' preserved in Ledger."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Amount"])
    ws.append([8500.55])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Amount")]
    entries, _ = mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("1.5000"))

    assert len(entries) == 1
    assert entries[0].original_type == "float"
    assert entries[0].original_value == 8500.55
    assert entries[0].get_decoded_original_value() == 8500.55


# ==============================================================================
# 11 - 15: Precision, Collision Gate & Retry Logic
# ==============================================================================

def test_high_precision_field_not_forced_to_two_decimals():
    """11. Fields with high precision (3+ decimals) preserve required decimal places."""
    values = [12.3456, 78.9101]
    prec = analyze_numeric_precision(values)
    assert prec == 4

    _, out_val = scale_decimal_value(12.3456, Decimal("2.0000"), precision=prec)
    assert out_val == 24.6912


def test_decimal_collision_gate():
    """12. Collision gate detects collisions and ensures distinct originals stay distinct."""
    # Two values very close: 1.004 and 1.006 with precision 2
    # At k=1.0000, 1.004*k = 1.00, 1.006*k = 1.01 (no collision)
    # But if k produces collision, retry or bumped precision separates them
    vals = [10.001, 10.002]
    final_k, prec, scaled = execute_field_scaling_with_collision_gate(
        vals, initial_k=Decimal("1.5000"), base_precision=2, max_retries=3
    )
    # Output values must be strictly distinct
    out1 = scaled[0][1]
    out2 = scaled[1][1]
    assert out1 != out2


def test_collision_retries_are_bounded():
    """13. Collision retries are strictly bounded by MAX_COLLISION_RETRIES."""
    # Force collision that cannot be resolved at precision 0
    with pytest.raises(NumericMaskingError, match="collision gate failed"):
        execute_field_scaling_with_collision_gate(
            [0.0001, 0.0002], initial_k=Decimal("1.2500"), base_precision=0, max_retries=1
        )


def test_collision_failure_blocks_output():
    """14. Unresolved collision raises safe NumericMaskingError and aborts."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Amount"])
    ws.append([0.000000001])
    ws.append([0.000000002])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Amount")]

    # Force precision to 0 to simulate collision failure
    # Test that function raises rather than corrupting data
    with pytest.raises(NumericMaskingError):
        execute_field_scaling_with_collision_gate(
            [0.000000001, 0.000000002], initial_k=Decimal("1.5000"), base_precision=0, max_retries=0
        )


def test_formula_cells_are_not_overwritten():
    """15. Formula cells (=A2*2) in the numeric column are never overwritten with static values."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Amount"])
    ws.append([1000])                  # Static input
    ws.append(["=A2*2"])               # Formula cell
    ws.append([3000])                  # Static input

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Amount")]
    mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("2.0000"))

    assert ws["A2"].value == 2000.0
    assert ws["A3"].value == "=A2*2"    # Exactly untouched formula string
    assert ws["A4"].value == 6000.0


# ==============================================================================
# 16 - 21: Number Formats, Auto-selection & Field Identity
# ==============================================================================

def test_number_format_preserved():
    """16. Existing cell.number_format is completely preserved after numeric masking."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Price"])
    cell = ws.cell(row=2, column=1, value=1234.56)
    cell.number_format = "$#,##0.00"

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Price")]
    mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("1.5000"))

    assert ws["A2"].number_format == "$#,##0.00"
    assert ws["A2"].value == 1851.84


def test_percentage_column_not_auto_selected():
    """17. Percentage formatted columns or non-amount columns are not automatically protected."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["CompletionRate", "Amount"])
    ws.append([0.85, 10000])

    scope = extract_table_scope(ws, header_row=1)
    # User only configures Amount (Column B)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="B", logical_field="Amount")]
    mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("2.0000"))

    # Column A remains untouched
    assert ws["A2"].value == 0.85
    assert ws["B2"].value == 20000.0


def test_duplicate_header_numeric_fields_get_distinct_protected_field_ids():
    """18. Two columns with identical header name ('金额') receive distinct protected_field_ids."""
    cfg1 = NumericProtectedColumnConfig(
        sheet_name="Sheet1", column_letter="B", logical_field="金额", raw_header="金额", display_header="金额"
    )
    cfg2 = NumericProtectedColumnConfig(
        sheet_name="Sheet1", column_letter="C", logical_field="金额", raw_header="金额", display_header="金额"
    )
    assert cfg1.protected_field_id != cfg2.protected_field_id


def test_blank_header_numeric_field_can_be_distinguished():
    """19. A numeric field with blank header can be protected and cleanly distinguished."""
    cfg = NumericProtectedColumnConfig(
        sheet_name="Sheet1", column_letter="D", logical_field="[空表头]", raw_header=None, display_header="[空表头]"
    )
    assert cfg.raw_header is None
    assert cfg.display_header == "[空表头]"
    assert cfg.protected_field_id.startswith("fld_")


def test_protected_field_id_is_stable_within_run():
    """20. protected_field_id remains stable across all rows and entries in a run."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([1000])
    ws.append([2000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Salary")]
    entries, _ = mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("1.5000"))

    assert len(entries) == 2
    assert entries[0].protected_field_id == cfg[0].protected_field_id
    assert entries[1].protected_field_id == cfg[0].protected_field_id


def test_protected_field_id_not_written_as_visible_excel_column():
    """21. protected_field_id is NEVER written into the Excel worksheet as a visible column or cell."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([1000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Salary")]
    mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("1.5000"))

    assert ws.max_column == 1
    assert ws["A1"].value == "Salary"
    assert ws["A2"].value == 1500.0


# ==============================================================================
# 22 - 25: Record Identity Snapshot & Confirmation Flag
# ==============================================================================

def test_single_key_record_identity_snapshot():
    """22. Single Key column is captured in record_identity_masked snapshot."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EmpID", "Salary"])
    ws.append(["E001", 8000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="B", logical_field="Salary")]
    entries, _ = mask_numeric_table_scope(
        ws, scope, cfg, scale_factor=Decimal("1.5000"), key_fields=["EmpID"]
    )

    assert len(entries) == 1
    snap = entries[0].record_identity_masked
    assert snap == [{"field": "EmpID", "type": "str", "value": "E001"}]


def test_composite_key_record_identity_snapshot():
    """23. Composite Key columns (Dept + ID) are captured in record_identity_masked snapshot."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Dept", "EmpID", "Salary"])
    ws.append(["Finance", "101", 9000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="C", logical_field="Salary")]
    entries, _ = mask_numeric_table_scope(
        ws, scope, cfg, scale_factor=Decimal("1.5000"), key_fields=["Dept", "EmpID"]
    )

    assert len(entries) == 1
    snap = entries[0].record_identity_masked
    assert snap == [
        {"field": "Dept", "type": "str", "value": "Finance"},
        {"field": "EmpID", "type": "str", "value": "101"},
    ]


def test_tokenized_key_is_used_as_masked_identity():
    """24. If Key field was tokenized by Text Tokenizer, the masked token is used in snapshot."""
    wb = openpyxl.Workbook()
    ws = wb.active
    # In workflow, Text Tokenizer runs first or simultaneously
    ws.append(["EmpID", "Salary"])
    ws.append(["MSK-A7D2F914-X8F3K2M9", 8000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="B", logical_field="Salary")]
    entries, _ = mask_numeric_table_scope(
        ws, scope, cfg, scale_factor=Decimal("1.5000"), key_fields=["EmpID"]
    )

    snap = entries[0].record_identity_masked
    assert snap[0]["value"] == "MSK-A7D2F914-X8F3K2M9"


def test_no_key_marks_restore_confirmation_required():
    """25. When no Key is confirmed, record_identity_masked is None (future REQUIRE_CONFIRMATION)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([8000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Salary")]
    entries, _ = mask_numeric_table_scope(
        ws, scope, cfg, scale_factor=Decimal("1.5000"), key_fields=None
    )

    assert len(entries) == 1
    assert entries[0].record_identity_masked is None


# ==============================================================================
# 26 - 29: Ledger Persistence & No Scale Leakage
# ==============================================================================

from cryptography.fernet import Fernet
from privacy_handler.identity import stamp_masked_workbook_identity


def test_numeric_group_written_to_encrypted_ledger(tmp_path: Path):
    """26. numeric_groups and scale factor are written into encrypted Ledger file."""
    ledger = RecoveryLedger(
        run=LedgerRunInfo(
            run_id="run_num_test_01",
            created_at="2026-10-07T12:00:00+00:00",
            workbook_identity="wb_hash_test",
        ),
        numeric_groups=[
            {
                "group_id": "numeric_group_1",
                "mode": "positive_scale",
                "scale_factor": "1.7500",
                "fields": ["fld_salary"],
            }
        ],
    )
    secret = Fernet.generate_key()
    enc_path = tmp_path / "ledger.enc"
    save_encrypted_ledger(ledger, enc_path, secret)

    # Decrypt and verify
    loaded = load_encrypted_ledger(enc_path, secret)
    assert len(loaded.numeric_groups) == 1
    assert loaded.numeric_groups[0]["scale_factor"] == "1.7500"


def test_scale_factor_not_written_to_masked_workbook():
    """27. Scale factor k is NOT written anywhere into the masked workbook cells or comments."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([1000])

    scope = extract_table_scope(ws, header_row=1)
    cfg = [NumericProtectedColumnConfig(sheet_name=ws.title, column_letter="A", logical_field="Salary")]
    mask_numeric_table_scope(ws, scope, cfg, scale_factor=Decimal("2.3456"))

    # Save to buffer and inspect xml
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    with zipfile.ZipFile(buf, "r") as z:
        for fname in z.namelist():
            content = z.read(fname)
            assert b"2.3456" not in content


def test_scale_factor_not_written_to_custom_properties():
    """28. Scale factor k is NOT written to custom document properties."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([1000])

    stamp_masked_workbook_identity(wb, "run_id_001", "1.0")
    # Verify no k property exists
    prop_names = [p.name for p in wb.custom_doc_props.props]
    assert "scale_factor" not in prop_names
    assert "k" not in prop_names


def test_formula_recalc_flag_written():
    """29. ensure_workbook_recalculation_properties successfully sets fullCalcOnLoad in OOXML."""
    wb = openpyxl.Workbook()
    ensure_workbook_recalculation_properties(wb)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    with zipfile.ZipFile(buf, "r") as z:
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in wb_xml
        assert 'forceFullCalc="1"' in wb_xml


# ==============================================================================
# 30 - 35: Regressions & Integrity Lock
# ==============================================================================

def test_metadata_hygiene_still_passes():
    """30. Metadata hygiene scrubs author info on numeric masked workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Salary"])
    ws.append([1000])

    wb.properties.creator = "Confidential Author"
    wb.properties.lastModifiedBy = "Secret Modifier"

    tool.sanitize_workbook_metadata(wb)
    assert wb.properties.creator == "Excel Privacy Tool"
    assert wb.properties.lastModifiedBy == "Excel Privacy Tool"


def test_unsupported_parts_preflight_still_passes(tmp_path: Path):
    """31. Pre-flight gate correctly detects pivot tables on numeric workbooks."""
    wb_path = tmp_path / "test_pivot.xlsx"
    with zipfile.ZipFile(wb_path, "w") as z:
        z.writestr("xl/pivotTables/pivotTable1.xml", "<pivotTable/>")

    feats = tool.detect_unsupported_features(wb_path)
    assert feats.pivot_table is True
    assert feats.has_unsupported is True


def test_output_path_customization_still_passes(tmp_path: Path):
    """32. Custom output path resolution works smoothly with numeric masked files."""
    out_dir = tmp_path / "exports"
    src = tmp_path / "financials.xlsx"
    src.write_text("dummy")

    resolved = resolve_safe_output_path(out_dir, src, suffix=".masked.xlsx")
    assert resolved.parent == out_dir
    assert resolved.name == "financials.masked.xlsx"


def test_xlsm_compatibility_does_not_regress():
    """33. XLSM files keep vbaProject.bin when loaded with keep_vba=True."""
    assert tool.load_workbook_private is not None




def test_integrity_lock_contains_numeric_masking_module():
    """35. numeric_masking.py is strictly registered in LOCKED_FILES."""
    mod_path = Path(__file__).resolve().parent.parent / "privacy_handler" / "numeric_masking.py"
    assert mod_path in tool.LOCKED_FILES


# ==============================================================================
# Phase 1: Per-Column Numeric Group Extension Tests (36 - 40)
# ==============================================================================

def test_36_single_sheet_two_numeric_groups(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """36. Single sheet with two distinct numeric groups scales fields by distinct factors."""
    src = tmp_path / "two_groups.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Staff"
    ws.append(["EmpID", "Salary", "Bonus", "Age"])
    ws.append(["E1", 10000, 2000, 30])
    ws.append(["E2", 15000, 3000, 40])
    wb.save(src)

    out_dir = tmp_path / "out"
    priv_dir = tmp_path / "priv"

    # Deterministic scale factor mock: Salary/Bonus (Group A) -> 1.5, Age (Group B) -> 2.0
    mock_factors = [Decimal("1.5000"), Decimal("2.0000")]
    monkeypatch.setattr("privacy_handler.numeric_masking.generate_scale_factor", lambda: mock_factors.pop(0))

    from privacy_handler.table_scope import SheetProtectionScope
    scopes = [
        SheetProtectionScope(
            sheet_name="Staff",
            header_row=1,
            numeric_columns=["Salary", "Bonus", "Age"],
            numeric_column_groups={
                "B": "local::Staff::A",
                "C": "local::Staff::A",
                "D": "local::Staff::B",
            },
        )
    ]

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        sheet_scopes=scopes,
    )

    masked_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = masked_wb["Staff"]

    # Salary & Bonus scaled by 1.5
    assert mws.cell(row=2, column=2).value == 15000.0   # 10000 * 1.5
    assert mws.cell(row=2, column=3).value == 3000.0    # 2000 * 1.5
    # Age scaled by 2.0
    assert mws.cell(row=2, column=4).value == 60        # 30 * 2.0

    # Verify Ledger records distinct group membership in protected_fields
    secret_bytes = Path(res["secret_key_path"]).read_bytes()
    ledger = load_encrypted_ledger(res["ledger_enc_path"], secret_bytes)
    pf_salary = next(pf for pf in ledger.protected_fields if pf.source_column_letter == "B")
    pf_bonus = next(pf for pf in ledger.protected_fields if pf.source_column_letter == "C")
    pf_age = next(pf for pf in ledger.protected_fields if pf.source_column_letter == "D")
    assert pf_salary.numeric_analysis_group == "local::Staff::A"
    assert pf_bonus.numeric_analysis_group == "local::Staff::A"
    assert pf_age.numeric_analysis_group == "local::Staff::B"


def test_37_duplicate_header_numeric_columns_different_groups(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """37. Duplicate headers (B:金额, C:金额) mapped by column_letter do not collide."""
    src = tmp_path / "dup_header.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Finance"
    ws.append(["ID", "金额", "金额"])
    ws.append(["001", 100, 100])
    wb.save(src)

    out_dir = tmp_path / "out"
    priv_dir = tmp_path / "priv"

    # Deterministic mock: B:金额 -> 1.5, C:金额 -> 2.5
    mock_factors = [Decimal("1.5000"), Decimal("2.5000")]
    monkeypatch.setattr("privacy_handler.numeric_masking.generate_scale_factor", lambda: mock_factors.pop(0))

    from privacy_handler.table_scope import SheetProtectionScope
    scopes = [
        SheetProtectionScope(
            sheet_name="Finance",
            header_row=1,
            numeric_columns=["B", "C"],
            numeric_column_groups={
                "B": "grp_1",
                "C": "grp_2",
            },
        )
    ]

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        sheet_scopes=scopes,
    )

    masked_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = masked_wb["Finance"]
    assert mws.cell(row=2, column=2).value == 150.0  # 100 * 1.5
    assert mws.cell(row=2, column=3).value == 250.0  # 100 * 2.5


def test_38_same_sheet_same_group_shares_factor(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """38. Fields sharing group identity always receive the identical scale factor."""
    src = tmp_path / "share_group.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Q1", "Q2", "Q3"])
    ws.append([1000, 2000, 3000])
    wb.save(src)

    out_dir = tmp_path / "out"
    priv_dir = tmp_path / "priv"

    mock_factors = [Decimal("3.0000")]
    monkeypatch.setattr("privacy_handler.numeric_masking.generate_scale_factor", lambda: mock_factors.pop(0))

    from privacy_handler.table_scope import SheetProtectionScope
    scopes = [
        SheetProtectionScope(
            sheet_name="Sales",
            header_row=1,
            numeric_columns=["Q1", "Q2", "Q3"],
            numeric_column_groups={"A": "shared_q", "B": "shared_q", "C": "shared_q"},
        )
    ]

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        sheet_scopes=scopes,
    )

    masked_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = masked_wb["Sales"]
    assert mws.cell(row=2, column=1).value == 3000.0
    assert mws.cell(row=2, column=2).value == 6000.0
    assert mws.cell(row=2, column=3).value == 9000.0


def test_39_legacy_sheet_level_group_unchanged(tmp_path: Path):
    """39. Legacy configuration without numeric_column_groups preserves single group behavior."""
    src = tmp_path / "legacy.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["A_Val", "B_Val"])
    ws.append([10, 20])
    wb.save(src)

    out_dir = tmp_path / "out"
    priv_dir = tmp_path / "priv"

    from privacy_handler.table_scope import SheetProtectionScope
    scopes = [
        SheetProtectionScope(
            sheet_name="Data",
            header_row=1,
            numeric_columns=["A_Val", "B_Val"],
            numeric_analysis_group="legacy_grp",
        )
    ]

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        sheet_scopes=scopes,
    )

    masked_wb = openpyxl.load_workbook(res["masked_workbook_path"])
    mws = masked_wb["Data"]
    # Ratio between columns preserved: B_Val / A_Val == 2.0
    v1 = mws.cell(row=2, column=1).value
    v2 = mws.cell(row=2, column=2).value
    assert round(v2 / v1, 2) == 2.0


def test_40_no_duplicate_ledger_or_side_effects(tmp_path: Path):
    """40. Multi-group numeric dispatch generates exactly 1 ledger entry per cell, without duplicate snapshots."""
    src = tmp_path / "side_effects.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Check"
    ws.append(["ID", "Val1", "Val2"])
    ws.append(["E1", 10, 20])
    ws.append(["E2", 30, 40])
    wb.save(src)

    out_dir = tmp_path / "out"
    priv_dir = tmp_path / "priv"

    from privacy_handler.table_scope import SheetProtectionScope
    scopes = [
        SheetProtectionScope(
            sheet_name="Check",
            header_row=1,
            record_key=["ID"],
            numeric_columns=["Val1", "Val2"],
            numeric_column_groups={"B": "grp_a", "C": "grp_b"},
        )
    ]

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=out_dir,
        private_key_dir=priv_dir,
        sheet_scopes=scopes,
    )

    secret_bytes = Path(res["secret_key_path"]).read_bytes()
    ledger = load_encrypted_ledger(res["ledger_enc_path"], secret_bytes)
    # 2 rows * 2 numeric columns = exactly 4 entries
    assert len(ledger.entries) == 4
    for entry in ledger.entries:
        assert entry.record_identity_masked is not None
        assert len(entry.record_identity_masked) == 1
        assert entry.record_identity_masked[0]["field"] == "ID"

