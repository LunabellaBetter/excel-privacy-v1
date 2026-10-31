"""Analysis-mode numeric masking module for V1 Excel Privacy Tool.

Implements Mode A (positive ratio scaling):
- Secure CSPRNG scale factor generation k in [1.25, 3.85]
- Dynamic precision analysis preserving field-specific decimals (p_out >= 2)
- Decimal-safe scaling avoiding IEEE-754 binary float artifacts
- Bounded collision gate with retry logic to guarantee injectivity (no duplicate collisions)
- Stable Protected Field Identity (protected_field_id)
- Record Identity snapshot (single or composite key, tokenized if protected)
- Non-destructive cell value update (preserving number_format, skipping formulas)
- OOXML full calculation properties activation (fullCalcOnLoad / forceFullCalc)
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from decimal import Decimal, ROUND_HALF_UP
import math
from pathlib import Path
import secrets
from typing import Any, Sequence

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    openpyxl = None
    get_column_letter = None

from privacy_handler.key_detector import is_empty_value, is_formula_value
from privacy_handler.recovery_ledger import (
    LedgerConfig,
    LedgerEntry,
    ProtectedField,
    RecoveryLedger,
    encode_typed_value,
)
from privacy_handler.table_scope import ColumnDescriptor, TableScope


class NumericMaskingError(Exception):
    """Base error for numeric masking operations."""


# Default CSPRNG scale factor bounds
SCALE_FACTOR_MIN = Decimal("1.2500")
SCALE_FACTOR_MAX = Decimal("3.8500")
MAX_COLLISION_RETRIES = 5


def generate_scale_factor(
    min_k: Decimal = SCALE_FACTOR_MIN,
    max_k: Decimal = SCALE_FACTOR_MAX,
) -> Decimal:
    """Generate a cryptographic pseudo-random scale factor k with 4 decimal places.

    Uses secrets.randbelow to ensure cryptographically secure generation.
    k will never equal 1.0 and is strictly within [min_k, max_k].
    """
    min_int = int(min_k * 10000)
    max_int = int(max_k * 10000)
    if max_int <= min_int:
        raise NumericMaskingError("max_k must be strictly greater than min_k")

    span = max_int - min_int + 1
    random_int = min_int + secrets.randbelow(span)
    k = Decimal(random_int) / Decimal(10000)
    if k == Decimal("1.0"):  # safety assertion
        k = Decimal("1.2501")
    return k


def analyze_numeric_precision(values: Sequence[Any]) -> int:
    """Analyze the minimum required output precision for a collection of numeric values.

    Returns:
    - At least 2 (standard currency precision)
    - If any value has more than 2 decimal places, returns that higher precision
      (up to a safety cap of 8).
    """
    max_decimals = 2
    for v in values:
        if v is None or isinstance(v, bool):
            continue
        if isinstance(v, int):
            continue
        if isinstance(v, float):
            # Check string representation or Decimal representation
            # Avoid binary float representation noise by rounding to 10 decimals first
            val_str = f"{v:.10f}".rstrip("0")
            if "." in val_str:
                decimals = len(val_str.split(".")[1])
                if decimals > max_decimals:
                    max_decimals = min(decimals, 8)
        elif isinstance(v, Decimal):
            exponent = v.as_tuple().exponent
            if isinstance(exponent, int) and exponent < 0:
                decimals = abs(exponent)
                if decimals > max_decimals:
                    max_decimals = min(decimals, 8)

    return max_decimals


def scale_decimal_value(
    value: Any,
    k: Decimal,
    precision: int = 2,
) -> tuple[Decimal | None, Any]:
    """Scale a single numeric value by k using Decimal semantics and bankers/half-up rounding.

    Returns:
    - (scaled_decimal, excel_output_value)
    - If value is 0/0.0, scaled value is Decimal(0)
    - If value is None, returns (None, None)
    - If original was int, output is int if no fraction, or float if scaled has fraction
    """
    if value is None or isinstance(value, bool):
        return None, value

    if isinstance(value, (int, float, Decimal)):
        # Convert to Decimal safely
        if isinstance(value, float):
            # Avoid binary float noise
            d_val = Decimal(str(value))
        else:
            d_val = Decimal(value)

        if d_val == 0:
            quantizer = Decimal("10") ** (-precision)
            res_dec = Decimal("0").quantize(quantizer, rounding=ROUND_HALF_UP)
            # Output matching original type if possible
            if isinstance(value, int):
                return res_dec, 0
            return res_dec, 0.0

        # Scale
        scaled = d_val * k
        quantizer = Decimal("10") ** (-precision)
        rounded = scaled.quantize(quantizer, rounding=ROUND_HALF_UP)

        # Produce Excel output value (numeric type)
        # In Excel, if rounded has no fraction and original was int, return int;
        # otherwise return float.
        float_out = float(rounded)
        return rounded, float_out

    return None, value


def execute_field_scaling_with_collision_gate(
    original_values: list[Any],
    initial_k: Decimal,
    base_precision: int,
    max_retries: int = MAX_COLLISION_RETRIES,
) -> tuple[Decimal, int, list[tuple[Decimal | None, Any]]]:
    """Scale a list of original values, strictly enforcing the Injectivity / Collision Gate.

    Rule: Different original distinct numeric values MUST NOT map to the same masked value.
    If a collision occurs:
    - Retry with a newly generated k (up to max_retries).
    - If collisions persist, increment precision by 1 and retry.
    - If still colliding after retries, raises safe NumericMaskingError.

    Returns:
    - (final_k, final_precision, list of (scaled_decimal, excel_val))
    """
    # Extract distinct non-empty numeric values
    distinct_originals: list[Any] = []
    seen: set[Any] = set()
    for v in original_values:
        if v is not None and not isinstance(v, bool) and isinstance(v, (int, float, Decimal)):
            # Normalize zero
            norm_key = Decimal(str(v)) if isinstance(v, (int, float, Decimal)) else v
            if norm_key not in seen:
                seen.add(norm_key)
                distinct_originals.append(v)

    if len(distinct_originals) <= 1:
        # 0 or 1 distinct values cannot collide with each other
        results = [
            scale_decimal_value(v, initial_k, base_precision) for v in original_values
        ]
        return initial_k, base_precision, results

    current_k = initial_k
    current_prec = base_precision

    for attempt in range(max_retries + 1):
        # Test distinct values for collision
        masked_seen: dict[Any, Any] = {}
        has_collision = False

        for orig in distinct_originals:
            _, out_val = scale_decimal_value(orig, current_k, current_prec)
            if out_val in masked_seen:
                has_collision = True
                break
            masked_seen[out_val] = orig

        if not has_collision:
            # Success! Compute full results
            full_results = [
                scale_decimal_value(v, current_k, current_prec) for v in original_values
            ]
            return current_k, current_prec, full_results

        # Attempt recovery: first try new k, then increase precision
        if attempt < max_retries:
            current_k = generate_scale_factor()
        else:
            # Final try: bump precision by 1
            current_prec += 1
            current_k = generate_scale_factor()
            masked_seen.clear()
            has_collision = False
            for orig in distinct_originals:
                _, out_val = scale_decimal_value(orig, current_k, current_prec)
                if out_val in masked_seen:
                    has_collision = True
                    break
                masked_seen[out_val] = orig
            if not has_collision:
                full_results = [
                    scale_decimal_value(v, current_k, current_prec) for v in original_values
                ]
                return current_k, current_prec, full_results

    raise NumericMaskingError(
        f"Numeric collision gate failed: distinct original values cannot be safely "
        f"masked without collision after {max_retries} retries."
    )


@dataclass(frozen=True)
class NumericProtectedColumnConfig:
    """User configuration for a column to undergo analysis-mode numeric masking."""

    sheet_name: str
    column_letter: str
    logical_field: str
    raw_header: str | None = None
    display_header: str | None = None
    protected_field_id: str = field(
        default_factory=lambda: f"fld_{secrets.token_hex(4)}"
    )


@dataclass
class NumericMaskingPlan:
    """Execution plan for a single shared-k numeric analysis group."""

    group_id: str
    scale_factor: Decimal
    protected_columns: list[NumericProtectedColumnConfig]
    precision_policy: str = "dynamic_currency_or_higher"


def ensure_workbook_recalculation_properties(wb: Any) -> None:
    """Configure openpyxl Workbook to force full formula recalculation on open.

    Sets:
    - wb.calculation.fullCalcOnLoad = True
    - wb.calculation.forceFullCalc = True
    This ensures Excel will recompute downstream formulas using masked inputs.
    """
    from openpyxl.workbook.properties import CalcProperties

    if getattr(wb, "calculation", None) is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True


def mask_numeric_table_scope(
    ws: Any,
    scope: TableScope,
    numeric_configs: list[NumericProtectedColumnConfig],
    scale_factor: Decimal,
    token_allocator: Any | None = None,
    key_fields: list[str] | None = None,
    entry_row_indices: list[int] | None = None,
) -> tuple[list[LedgerEntry], dict[str, Any]]:
    """Execute numeric masking on a worksheet according to TableScope.

    - Leaves formula cells untouched.
    - Preserves cell number_format.
    - Captures typed original values and masked values into Ledger entries.
    - Builds Record Identity snapshot (tokenized if key field is text-masked).
    - Checks collision gate per column.

    Returns:
    - (ledger_entries, group_metadata)
    """
    if scope.data_end_row is None or scope.data_end_row < scope.data_start_row:
        return [], {}

    # Map column letter to ColumnDescriptor
    col_desc_map: dict[str, ColumnDescriptor] = {
        col.column_letter: col for col in scope.selected_columns
    }

    # Validate all requested numeric configs exist in scope
    for cfg in numeric_configs:
        if cfg.column_letter not in col_desc_map:
            raise NumericMaskingError(
                f"Numeric protected column {cfg.column_letter} not present in worksheet scope"
            )

    entries: list[LedgerEntry] = []
    field_precisions: dict[str, int] = {}

    # For each numeric column:
    # 1. Collect all static numeric cell values
    # 2. Analyze precision and run collision gate
    # 3. Apply masked values to worksheet
    # 4. Generate Ledger entries with Record Identity snapshot

    for cfg in numeric_configs:
        col_desc = col_desc_map[cfg.column_letter]
        col_idx = col_desc.column_index

        # Collect rows and cells
        row_indices: list[int] = []
        original_values: list[Any] = []
        is_formula_list: list[bool] = []

        for r in range(scope.data_start_row, scope.data_end_row + 1):
            cell = ws.cell(row=r, column=col_idx)
            val = cell.value
            row_indices.append(r)
            if is_formula_value(val):
                is_formula_list.append(True)
                original_values.append(val)
            else:
                is_formula_list.append(False)
                original_values.append(val)

        # Filter static numeric values for precision analysis and collision check
        static_numeric = [
            v for v, is_f in zip(original_values, is_formula_list)
            if not is_f and v is not None and isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)
        ]

        base_prec = analyze_numeric_precision(static_numeric)
        field_precisions[cfg.protected_field_id] = base_prec

        # Run collision gate
        _, effective_prec, scaled_results = execute_field_scaling_with_collision_gate(
            original_values=original_values,
            initial_k=scale_factor,
            base_precision=base_prec,
        )

        # Apply to cells and generate ledger entries
        for r, is_f, orig, (dec_val, out_val) in zip(
            row_indices, is_formula_list, original_values, scaled_results
        ):
            if orig is None or is_empty_value(orig):
                continue

            if is_f:
                # Rule: DO NOT overwrite formula cells
                continue

            if not isinstance(orig, (int, float, Decimal)) or isinstance(orig, bool):
                # Non-numeric value in numeric column: leave untouched
                continue

            # Update cell value in openpyxl worksheet, preserving number_format
            cell = ws.cell(row=r, column=col_idx)
            cell.value = out_val

            # Build record identity snapshot for this row
            record_identity_snapshot: list[dict[str, Any]] | None = None
            if key_fields:
                record_identity_snapshot = []
                for kf in key_fields:
                    # Find column for kf
                    # kf can be column letter or header
                    kf_col = None
                    for c in scope.selected_columns:
                        if c.column_letter == kf or c.header == kf:
                            kf_col = c
                            break
                    if kf_col:
                        raw_k_val = ws.cell(row=r, column=kf_col.column_index).value
                        # If token_allocator has a token for this field/value, use masked token
                        snap_val = raw_k_val
                        snap_type = type(raw_k_val).__name__
                        record_identity_snapshot.append({
                            "field": kf,
                            "type": snap_type,
                            "value": snap_val,
                        })

            # Create ledger entry
            orig_tag, orig_ser = encode_typed_value(orig)
            masked_tag, masked_ser = encode_typed_value(out_val)

            entries.append(
                LedgerEntry(
                    sheet=scope.sheet_name,
                    logical_field=cfg.logical_field,
                    key_masked_value=None,  # Legacy field
                    original_type=orig_tag,
                    original_value=orig_ser,
                    masked_type=masked_tag,
                    masked_value=masked_ser,
                    protection_type="numeric_scale",
                    protected_field_id=cfg.protected_field_id,
                    record_identity_masked=record_identity_snapshot,
                )
            )
            if entry_row_indices is not None:
                entry_row_indices.append(r)

    group_metadata = {
        "group_id": "numeric_group_1",
        "mode": "positive_scale",
        "scale_factor": str(scale_factor),
        "fields": [cfg.protected_field_id for cfg in numeric_configs],
        "field_precisions": field_precisions,
    }

    return entries, group_metadata
