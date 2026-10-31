"""Conflict & Compare Engine for V1 AI round-trip privacy protocol.

Compares an AI-returned workbook against an encrypted Recovery Ledger.
Detects:
- Row-level structural changes (ADDED_ROW, DELETED_ROW, whole-row sorts preserved as UNCHANGED)
- Column-level structural changes (ADDED_COLUMN, DELETED_COLUMN, MOVED_OR_RELOCATED_COLUMN)
- Cell modifications (MODIFIED, UNCHANGED)
- Critical security violations:
  * PROTECTED_VALUE_CONFLICT (tampered, swapped, or moved protected text/number across records)
  * RECORD_IDENTITY_CONFLICT (duplicate or corrupted primary key in returned workbook)
  * REQUIRE_CONFIRMATION (missing record identity, ambiguous returned columns, or unresolved mappings)
- Diagnostic Privacy:
  * Redacted / hashed representation of Record Identity in default exports / to_dict()
  * Zero original sensitive values in comparison objects or logs.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation
from enum import Enum
import hashlib
from pathlib import Path
from typing import Any, Sequence

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    openpyxl = None
    get_column_letter = None

from privacy_handler.key_detector import is_empty_value, is_formula_value
from privacy_handler.recovery_ledger import (
    LedgerEntry,
    ProtectedField,
    RecoveryLedger,
    decode_typed_value,
    load_encrypted_ledger,
)
from privacy_handler.table_scope import extract_table_scope


class ConflictEngineError(Exception):
    """Base error for conflict and compare engine operations."""


class ChangeStatus(str, Enum):
    UNCHANGED = "UNCHANGED"
    MODIFIED = "MODIFIED"
    ADDED_ROW = "ADDED_ROW"
    DELETED_ROW = "DELETED_ROW"
    ADDED_COLUMN = "ADDED_COLUMN"
    DELETED_COLUMN = "DELETED_COLUMN"
    MOVED_OR_RELOCATED_COLUMN = "MOVED_OR_RELOCATED_COLUMN"
    PROTECTED_VALUE_CONFLICT = "PROTECTED_VALUE_CONFLICT"
    RECORD_IDENTITY_CONFLICT = "RECORD_IDENTITY_CONFLICT"
    REQUIRE_CONFIRMATION = "REQUIRE_CONFIRMATION"


def _format_record_key_redacted(record_key: Any, alias_map: dict[str, str] | None = None) -> str:
    """Produce a privacy-safe report/session alias for a Record Key.

    - Tokens (e.g. MSK-A7D2F914-X8F3K2M9) and row keys (row_1) are safe to show as-is.
    - Raw plaintext values are replaced with ephemeral report aliases (Record-01, Record-02)
      or session-scoped safe tokens, never exposing plain SHA256 of sensitive values.
    """
    if record_key is None:
        return "record #UNKNOWN"

    s_key = str(record_key)
    if s_key.startswith("MSK-") or s_key.startswith("row_"):
        return s_key

    def _get_alias(val_str: str) -> str:
        if val_str.startswith("MSK-") or val_str.startswith("row_"):
            return val_str
        if alias_map is not None:
            if val_str not in alias_map:
                alias_map[val_str] = f"Record-{len(alias_map) + 1:02d}"
            return alias_map[val_str]
        # Standalone fallback when no alias_map is passed:
        # Use ephemeral safe label format without plain deterministic SHA256
        return "Record-01"

    # If it's a tuple of (field, val) or simple items
    if isinstance(record_key, tuple):
        components = []
        for item in record_key:
            if isinstance(item, tuple) and len(item) == 2:
                f_name, f_val = item
                s_val = str(f_val)
                components.append(f"{f_name}={_get_alias(s_val)}")
            else:
                s_item = str(item)
                components.append(_get_alias(s_item))
        return ", ".join(components)

    return _get_alias(s_key)


@dataclass
class CellChange:
    sheet: str
    row_idx: int | None
    col_idx: int | None
    col_letter: str | None
    logical_field: str
    protected_field_id: str | None
    protection_type: str | None  # "text_token", "numeric_scaled", or None (unprotected)
    record_key: Any              # Internal Record Key representation
    before_value: Any            # Expected masked value (never original real value)
    after_value: Any             # Returned value in AI workbook
    status: ChangeStatus
    notes: str = ""

    @property
    def record_key_display(self) -> str:
        """Safe redacted / tokenized Record Key representation for display or logs."""
        return _format_record_key_redacted(self.record_key)

    def to_dict(self, redact_keys: bool = True, alias_map: dict[str, str] | None = None) -> dict[str, Any]:
        """Serialize cell change. Defaults to redact_keys=True for safe sharing."""
        if redact_keys:
            key_repr = _format_record_key_redacted(self.record_key, alias_map=alias_map)
        else:
            key_repr = str(self.record_key) if self.record_key is not None else None
        return {
            "sheet": self.sheet,
            "row_idx": self.row_idx,
            "col_idx": self.col_idx,
            "col_letter": self.col_letter,
            "logical_field": self.logical_field,
            "protected_field_id": self.protected_field_id,
            "protection_type": self.protection_type,
            "record_key": key_repr,
            "before_value": self.before_value,
            "after_value": self.after_value,
            "status": self.status.value,
            "notes": self.notes,
        }

    def to_safe_dict(self, alias_map: dict[str, str] | None = None) -> dict[str, Any]:
        """Opt-in safe dictionary with redacted Record Identity."""
        return self.to_dict(redact_keys=True, alias_map=alias_map)

    def to_internal_dict(self) -> dict[str, Any]:
        """Internal serialization retaining raw in-memory record_key."""
        return self.to_dict(redact_keys=False)


@dataclass
class StructuralChange:
    change_type: ChangeStatus  # ADDED_ROW, DELETED_ROW, ADDED_COLUMN, DELETED_COLUMN, MOVED_OR_RELOCATED_COLUMN, RECORD_IDENTITY_CONFLICT
    sheet: str
    identifier: str  # Row key or column header
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self, redact_keys: bool = False, alias_map: dict[str, str] | None = None) -> dict[str, Any]:
        """Serialize structural change. If redact_keys=True, replaces row keys with safe aliases."""
        ident = self.identifier
        if redact_keys and (self.change_type in (ChangeStatus.ADDED_ROW, ChangeStatus.DELETED_ROW, ChangeStatus.RECORD_IDENTITY_CONFLICT)):
            # If identifier contains raw key values, redact them using alias_map
            ident = _format_record_key_redacted(ident, alias_map=alias_map)

        return {
            "change_type": self.change_type.value,
            "sheet": self.sheet,
            "identifier": ident,
            "details": self.details,
        }

    def to_safe_dict(self, alias_map: dict[str, str] | None = None) -> dict[str, Any]:
        """Opt-in safe dictionary with redacted identifiers."""
        return self.to_dict(redact_keys=True, alias_map=alias_map)

    def to_internal_dict(self) -> dict[str, Any]:
        """Internal memory serialization retaining raw identifiers."""
        return self.to_dict(redact_keys=False)


@dataclass
class ChangeList:
    cell_changes: list[CellChange] = field(default_factory=list)
    structural_changes: list[StructuralChange] = field(default_factory=list)
    has_conflicts: bool = False
    requires_confirmation: bool = False
    summary: dict[str, int] = field(default_factory=dict)
    matched_columns: dict[str, dict[str, int]] = field(default_factory=dict)  # sheet -> {logical_field: col_idx}
    matched_columns_by_pf_id: dict[str, dict[str, int]] = field(default_factory=dict)  # sheet -> {protected_field_id: col_idx}
    unmatched_columns: dict[str, list[str]] = field(default_factory=dict)     # sheet -> [header_names]

    def compute_summary(self) -> None:
        counts: dict[str, int] = {}
        for c in self.cell_changes:
            counts[c.status.value] = counts.get(c.status.value, 0) + 1
        for s in self.structural_changes:
            counts[s.change_type.value] = counts.get(s.change_type.value, 0) + 1

        self.summary = counts
        self.has_conflicts = (
            counts.get(ChangeStatus.PROTECTED_VALUE_CONFLICT.value, 0) > 0
            or counts.get(ChangeStatus.RECORD_IDENTITY_CONFLICT.value, 0) > 0
        )
        self.requires_confirmation = (
            counts.get(ChangeStatus.REQUIRE_CONFIRMATION.value, 0) > 0
            or self.has_conflicts
        )

    def to_dict(self, redact_keys: bool = True) -> dict[str, Any]:
        """Serialize ChangeList. Defaults to redact_keys=True for Diagnostic Privacy.

        Maintains an ephemeral safe_alias_map across all cell and structural changes
        so that identical raw record keys receive consistent aliases (Record-01, Record-02)
        within this report without deriving from plain SHA256 of sensitive keys.
        """
        self.compute_summary()
        alias_map: dict[str, str] = {} if redact_keys else None
        return {
            "has_conflicts": self.has_conflicts,
            "requires_confirmation": self.requires_confirmation,
            "summary": self.summary,
            "cell_changes": [c.to_dict(redact_keys=redact_keys, alias_map=alias_map) for c in self.cell_changes],
            "structural_changes": [s.to_dict(redact_keys=redact_keys, alias_map=alias_map) for s in self.structural_changes],
            "matched_columns": self.matched_columns,
            "unmatched_columns": self.unmatched_columns,
        }

    def to_safe_dict(self) -> dict[str, Any]:
        """Explicit opt-in safe dictionary with redacted Record Identity."""
        return self.to_dict(redact_keys=True)

    def to_internal_dict(self) -> dict[str, Any]:
        """Internal memory serialization retaining raw record keys."""
        return self.to_dict(redact_keys=False)


def _normalize_cell_scalar(val: Any) -> Any:
    """Normalize cell value for comparison without loss of type identity."""
    if val is None:
        return None
    if isinstance(val, str):
        stripped = val.strip()
        if not stripped:
            return None
        return stripped
    return val


def _values_are_equivalent(expected: Any, returned: Any, is_numeric: bool = False) -> bool:
    """Check value equivalence.

    For numeric values, compares Decimal representations to handle float representation noise.
    For non-numeric values, compares exact type and value or normalized string equality.
    """
    if expected is None and returned is None:
        return True
    if expected is None or returned is None:
        return False

    if is_numeric:
        try:
            d_exp = Decimal(str(expected)).normalize()
            d_ret = Decimal(str(returned)).normalize()
            return d_exp == d_ret
        except (InvalidOperation, TypeError, ValueError):
            return str(expected).strip() == str(returned).strip()

    # Exact equality
    if expected == returned:
        return True

    # If both can be string-compared
    if str(expected).strip() == str(returned).strip():
        return True

    return False


def _make_record_key_tuple(snap: list[dict[str, Any]] | None) -> tuple | None:
    """Convert record_identity_masked snapshot list into a hashable canonical tuple."""
    if not snap:
        return None
    components = []
    for item in snap:
        field_name = item.get("field", "")
        val = item.get("value")
        norm_val = str(val).strip() if val is not None else ""
        components.append((field_name, norm_val))
    return tuple(components)


def compare_workbook_against_ledger(
    returned_workbook_or_path: Any,
    ledger_or_path: RecoveryLedger | Path | str,
    secret_key: Any = None,
    target_sheet: str | None = None,
) -> ChangeList:
    """Compare an AI-returned workbook against the ground-truth Recovery Ledger.

    Produces a comprehensive ChangeList detailing structural, cell-level, and conflict changes.
    """
    # 1. Resolve Ledger
    if isinstance(ledger_or_path, RecoveryLedger):
        ledger = ledger_or_path
    else:
        ledger_path = Path(ledger_or_path).resolve()
        if secret_key is None:
            raise ConflictEngineError("secret_key is required when loading encrypted ledger from file path")
        ledger = load_encrypted_ledger(ledger_path, secret_key)

    # 2. Resolve Returned Workbook
    close_needed = False
    if isinstance(returned_workbook_or_path, (str, Path)):
        import openpyxl
        wb = openpyxl.load_workbook(returned_workbook_or_path, data_only=False)
        close_needed = True
    else:
        wb = returned_workbook_or_path

    change_list = ChangeList()

    try:
        # Determine sheets to compare
        sheets_in_ledger = {pf.sheet_name for pf in ledger.protected_fields if pf.sheet_name}
        if not sheets_in_ledger:
            sheets_in_ledger = {e.sheet for e in ledger.entries if e.sheet}

        if ledger.configuration and ledger.configuration.sheet_scopes:
            for sc in ledger.configuration.sheet_scopes:
                if "sheet_name" in sc:
                    sheets_in_ledger.add(sc["sheet_name"])

        if target_sheet:
            sheets_to_process = [target_sheet]
        elif sheets_in_ledger:
            sheets_to_process = sorted(list(sheets_in_ledger))
        else:
            sheets_to_process = list(wb.sheetnames)

        # Build protected fields lookup:
        # (sheet, protected_field_id) -> ProtectedField
        # (sheet, logical_field) -> ProtectedField (fallback)
        pf_by_id: dict[tuple[str, str], ProtectedField] = {}
        pf_by_field: dict[tuple[str, str], ProtectedField] = {}
        for pf in ledger.protected_fields:
            s_name = pf.sheet_name or ""
            pf_key = pf.protected_field_id or pf.logical_field
            pf_by_id[(s_name, pf_key)] = pf
            pf_by_field[(s_name, pf.logical_field)] = pf
            if pf.display_header:
                pf_by_field[(s_name, pf.display_header)] = pf
            if pf.raw_header:
                pf_by_field[(s_name, pf.raw_header)] = pf

        # Build ledger entries index:
        ledger_entries_by_key: dict[tuple[str, str, Any], LedgerEntry] = {}
        ledger_entries_by_sheet_field: dict[tuple[str, str], list[LedgerEntry]] = {}
        ledger_entries_by_sheet_pf_id: dict[tuple[str, str], list[LedgerEntry]] = {}
        expected_masked_pool: dict[tuple[str, str], set[str]] = {}

        for entry in ledger.entries:
            sh = entry.sheet
            lf = entry.logical_field
            pf_id = entry.protected_field_id or lf
            key_tup = _make_record_key_tuple(entry.record_identity_masked)
            if key_tup is not None:
                ledger_entries_by_key[(sh, pf_id, key_tup)] = entry
            ledger_entries_by_sheet_field.setdefault((sh, lf), []).append(entry)
            if entry.protected_field_id:
                ledger_entries_by_sheet_pf_id.setdefault((sh, entry.protected_field_id), []).append(entry)

            dec_masked = entry.get_decoded_masked_value()
            if dec_masked is not None:
                pool_set = expected_masked_pool.setdefault((sh, pf_id), set())
                expected_masked_pool.setdefault((sh, lf), pool_set)
                s_raw = str(dec_masked).strip()
                pool_set.add(s_raw)
                try:
                    d_val = Decimal(s_raw).normalize()
                    pool_set.add(str(d_val))
                    if d_val == d_val.to_integral():
                        pool_set.add(str(int(d_val)))
                except Exception:
                    pass

        confirmed_key_field = ledger.configuration.confirmed_key_field if ledger.configuration else None

        for sheet_name in sheets_to_process:
            if sheet_name not in wb.sheetnames:
                change_list.structural_changes.append(
                    StructuralChange(
                        change_type=ChangeStatus.DELETED_ROW,
                        sheet=sheet_name,
                        identifier=f"Sheet '{sheet_name}' is missing",
                    )
                )
                continue

            ws = wb[sheet_name]

            cfg_scope = None
            if ledger.configuration and ledger.configuration.sheet_scopes:
                for sc in ledger.configuration.sheet_scopes:
                    if sc.get("sheet_name") == sheet_name:
                        cfg_scope = sc
                        break

            header_row = cfg_scope.get("header_row", 1) if cfg_scope else 1
            data_start_row = cfg_scope.get("data_start_row", header_row + 1) if cfg_scope else header_row + 1

            returned_scope = extract_table_scope(ws, header_row=header_row, data_start_row=data_start_row)

            change_list.matched_columns[sheet_name] = {}
            change_list.matched_columns_by_pf_id[sheet_name] = {}
            change_list.unmatched_columns[sheet_name] = []

            # ------------------------------------------------------------------
            # Phase 1: Determine Key Columns in returned worksheet
            # ------------------------------------------------------------------
            sheet_entries = [e for e in ledger.entries if e.sheet == sheet_name]
            has_record_identity = False
            key_field_names: list[str] = []
            if sheet_entries and sheet_entries[0].record_identity_masked:
                has_record_identity = True
                key_field_names = [item["field"] for item in sheet_entries[0].record_identity_masked]
            elif confirmed_key_field:
                has_record_identity = True
                key_field_names = [f.strip() for f in confirmed_key_field.split("+")]

            returned_col_by_header: dict[str, int] = {}
            for col in returned_scope.selected_columns:
                returned_col_by_header.setdefault(col.header, col.column_index)
                if col.raw_header:
                    returned_col_by_header.setdefault(col.raw_header, col.column_index)

            key_col_indices: list[int] = []
            if has_record_identity:
                for kf in key_field_names:
                    idx = None
                    if kf in returned_col_by_header:
                        idx = returned_col_by_header[kf]
                    else:
                        for col in returned_scope.selected_columns:
                            if col.column_letter == kf:
                                idx = col.column_index
                                break
                    if idx is not None:
                        key_col_indices.append(idx)
                    else:
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.RECORD_IDENTITY_CONFLICT,
                                sheet=sheet_name,
                                identifier=f"Missing key field '{kf}' in returned sheet",
                            )
                        )
                        has_record_identity = False
                        break

            # Map returned rows by key tuple:
            # key_tuple -> list of row indices
            returned_rows_by_key: dict[tuple, list[int]] = {}
            if has_record_identity and returned_scope.data_end_row is not None and returned_scope.data_end_row >= returned_scope.data_start_row:
                for r in range(returned_scope.data_start_row, returned_scope.data_end_row + 1):
                    components = []
                    for kf, c_idx in zip(key_field_names, key_col_indices):
                        val = ws.cell(row=r, column=c_idx).value
                        norm_val = str(val).strip() if val is not None else ""
                        components.append((kf, norm_val))
                    kt = tuple(components)
                    returned_rows_by_key.setdefault(kt, []).append(r)

            # Check for duplicate keys in returned workbook -> RECORD_IDENTITY_CONFLICT
            if has_record_identity:
                for kt, row_list in returned_rows_by_key.items():
                    if len(row_list) > 1:
                        key_str = ", ".join(f"{f}={v}" for f, v in kt)
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.RECORD_IDENTITY_CONFLICT,
                                sheet=sheet_name,
                                identifier=key_str,
                                details={"duplicate_row_indices": row_list},
                            )
                        )

            # ------------------------------------------------------------------
            # Phase 2: Returned Column Resolution (Candidate Discovery -> Verification -> Unique Resolution)
            # ------------------------------------------------------------------
            sheet_pfs = [pf for (s, _), pf in pf_by_id.items() if s == sheet_name or s == ""]
            seen_pf_ids = set()
            unique_sheet_pfs: list[ProtectedField] = []
            for pf in sheet_pfs:
                pf_id = pf.protected_field_id or pf.logical_field
                if pf_id not in seen_pf_ids:
                    seen_pf_ids.add(pf_id)
                    unique_sheet_pfs.append(pf)

            matched_col_indices: set[int] = set()

            for pf in unique_sheet_pfs:
                pf_id = pf.protected_field_id
                # Get expected entries for this specific protected field
                expected_entries = ledger_entries_by_sheet_pf_id.get((sheet_name, pf_id)) if pf_id else None
                if not expected_entries:
                    expected_entries = ledger_entries_by_sheet_field.get((sheet_name, pf.logical_field), [])

                is_num = pf.protection_type == "numeric_scaled"
                expected_pool = expected_masked_pool.get((sheet_name, pf.logical_field), set())

                # Step 1: Candidate Discovery
                candidate_cols: list[Any] = []
                for col in returned_scope.selected_columns:
                    # Match by logical header, raw header, or display header
                    if (
                        col.header == pf.logical_field
                        or (pf.display_header and col.header == pf.display_header)
                        or (pf.raw_header and col.raw_header == pf.raw_header)
                    ):
                        candidate_cols.append(col)

                # Step 2: Evidence Scoring / Verification
                resolved_col = None

                if len(candidate_cols) == 1:
                    # Single candidate with matching header
                    cand = candidate_cols[0]
                    # Verify snapshot pattern if Record Identity exists
                    if has_record_identity and expected_entries:
                        match_count = 0
                        checked_count = 0
                        for entry in expected_entries:
                            kt = _make_record_key_tuple(entry.record_identity_masked)
                            if kt in returned_rows_by_key and len(returned_rows_by_key[kt]) == 1:
                                r_idx = returned_rows_by_key[kt][0]
                                actual_val = ws.cell(row=r_idx, column=cand.column_index).value
                                exp_val = entry.get_decoded_masked_value()
                                if _values_are_equivalent(exp_val, actual_val, is_numeric=is_num):
                                    match_count += 1
                                checked_count += 1
                        # When snapshot entries were checked:
                        # If a column with the expected header has 0 matches across all checked records,
                        # check whether any values in that column belong to the expected masked pool for this field.
                        # - If values belong to the pool (e.g. values swapped across records), bind candidate column
                        #   so cell-level conflict engine can report PROTECTED_VALUE_CONFLICT / SWAP.
                        # - If values do NOT belong to the pool at all (e.g. completely alien values, or new column
                        #   created after original was renamed/deleted), DO NOT bind! Downgrade to REQUIRE_CONFIRMATION.
                        if checked_count > 0 and match_count == 0:
                            has_pool_match = False
                            for entry in expected_entries:
                                kt = _make_record_key_tuple(entry.record_identity_masked)
                                if kt in returned_rows_by_key and len(returned_rows_by_key[kt]) == 1:
                                    r_idx = returned_rows_by_key[kt][0]
                                    actual_val = ws.cell(row=r_idx, column=cand.column_index).value
                                    if actual_val is not None:
                                        act_str = str(actual_val).strip()
                                        if act_str in expected_pool:
                                            has_pool_match = True
                                            break
                            if not has_pool_match:
                                change_list.structural_changes.append(
                                    StructuralChange(
                                        change_type=ChangeStatus.REQUIRE_CONFIRMATION,
                                        sheet=sheet_name,
                                        identifier=pf.logical_field,
                                        details={
                                            "reason": f"Single candidate column '{cand.column_letter}' failed snapshot verification for '{pf.logical_field}' (0/{checked_count} matched)",
                                            "candidate_columns": [cand.column_letter],
                                        },
                                    )
                                )
                                resolved_col = None
                            else:
                                resolved_col = cand
                        else:
                            resolved_col = cand
                    else:
                        resolved_col = cand

                elif len(candidate_cols) > 1:
                    # MULTIPLE CANDIDATES WITH SAME HEADER: Disambiguation required!
                    valid_snapshot_candidates = []
                    if has_record_identity and expected_entries:
                        for cand in candidate_cols:
                            match_count = 0
                            checked_count = 0
                            for entry in expected_entries:
                                kt = _make_record_key_tuple(entry.record_identity_masked)
                                if kt in returned_rows_by_key and len(returned_rows_by_key[kt]) == 1:
                                    r_idx = returned_rows_by_key[kt][0]
                                    actual_val = ws.cell(row=r_idx, column=cand.column_index).value
                                    exp_val = entry.get_decoded_masked_value()
                                    if _values_are_equivalent(exp_val, actual_val, is_numeric=is_num):
                                        match_count += 1
                                    checked_count += 1
                            if checked_count > 0 and match_count == checked_count:
                                valid_snapshot_candidates.append(cand)

                    if len(valid_snapshot_candidates) == 1:
                        # Exactly one candidate matched expected snapshot!
                        resolved_col = valid_snapshot_candidates[0]
                    else:
                        # Ambiguous: 0 or >1 candidates match pattern -> REQUIRE_CONFIRMATION!
                        # source_column_letter alone is forbidden from automatically guessing
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.REQUIRE_CONFIRMATION,
                                sheet=sheet_name,
                                identifier=pf.logical_field,
                                details={
                                    "reason": f"Ambiguous returned columns for '{pf.logical_field}': {len(candidate_cols)} candidate columns found",
                                    "candidate_columns": [c.column_letter for c in candidate_cols],
                                },
                            )
                        )
                        resolved_col = None

                else:
                    # candidate_cols == 0: No header matched
                    resolved_col = None

                # Step 3: Record Resolution Result
                if resolved_col is not None:
                    matched_idx = resolved_col.column_index
                    change_list.matched_columns[sheet_name][pf.logical_field] = matched_idx
                    pf_key = pf.protected_field_id or pf.logical_field
                    change_list.matched_columns_by_pf_id[sheet_name][pf_key] = matched_idx
                    matched_col_indices.add(matched_idx)
                    # Check if column was relocated
                    if pf.source_column_letter:
                        returned_letter = get_column_letter(matched_idx) if get_column_letter else str(matched_idx)
                        if returned_letter != pf.source_column_letter:
                            change_list.structural_changes.append(
                                StructuralChange(
                                    change_type=ChangeStatus.MOVED_OR_RELOCATED_COLUMN,
                                    sheet=sheet_name,
                                    identifier=pf.logical_field,
                                    details={
                                        "original_column": pf.source_column_letter,
                                        "returned_column": returned_letter,
                                        "col_index": matched_idx,
                                        "protected_field_id": pf.protected_field_id,
                                    },
                                )
                            )
                else:
                    # Column not resolved
                    change_list.unmatched_columns[sheet_name].append(pf.logical_field)
                    if len(candidate_cols) == 0:
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.DELETED_COLUMN,
                                sheet=sheet_name,
                                identifier=pf.logical_field,
                                details={"expected_column": pf.source_column_letter},
                            )
                        )

            # Check for ADDED_COLUMN in returned worksheet
            known_headers = {pf.logical_field for pf in unique_sheet_pfs}
            known_headers.update(pf.display_header for pf in unique_sheet_pfs if pf.display_header)
            if confirmed_key_field:
                known_headers.update(k.strip() for k in confirmed_key_field.split("+"))
            if cfg_scope and "selected_columns" in cfg_scope:
                for sc_col in cfg_scope["selected_columns"]:
                    if isinstance(sc_col, dict):
                        h = sc_col.get("header")
                        if h:
                            known_headers.add(h)

            for col in returned_scope.selected_columns:
                if col.column_index not in matched_col_indices:
                    if col.header not in known_headers:
                        change_list.unmatched_columns[sheet_name].append(col.header)
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.ADDED_COLUMN,
                                sheet=sheet_name,
                                identifier=col.header,
                                details={"column_letter": col.column_letter, "col_index": col.column_index},
                            )
                        )

            # ------------------------------------------------------------------
            # Phase 3: Row Mapping & Structural Changes
            # ------------------------------------------------------------------
            expected_keys_set: set[tuple] = set()
            for (sh, _, kt) in ledger_entries_by_key.keys():
                if sh == sheet_name and kt is not None:
                    expected_keys_set.add(kt)

            if has_record_identity:
                for kt in returned_rows_by_key.keys():
                    if kt not in expected_keys_set:
                        key_str = ", ".join(f"{f}={v}" for f, v in kt)
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.ADDED_ROW,
                                sheet=sheet_name,
                                identifier=key_str,
                                details={"row_indices": returned_rows_by_key[kt]},
                            )
                        )

                returned_keys_set = set(returned_rows_by_key.keys())
                for kt in expected_keys_set:
                    if kt not in returned_keys_set:
                        key_str = ", ".join(f"{f}={v}" for f, v in kt)
                        change_list.structural_changes.append(
                            StructuralChange(
                                change_type=ChangeStatus.DELETED_ROW,
                                sheet=sheet_name,
                                identifier=key_str,
                            )
                        )

            # ------------------------------------------------------------------
            # Phase 4: Cell-by-cell Comparison
            # ------------------------------------------------------------------
            # Case A: Record Identity Available
            if has_record_identity:
                for (sh, pf_id, kt), entry in ledger_entries_by_key.items():
                    if sh != sheet_name:
                        continue
                    
                    # Match column index: prefer protected_field_id, fallback to logical_field
                    col_idx = change_list.matched_columns_by_pf_id.get(sheet_name, {}).get(pf_id)
                    if col_idx is None:
                        col_idx = change_list.matched_columns.get(sheet_name, {}).get(entry.logical_field)
                    if col_idx is None:
                        continue

                    lf = entry.logical_field
                    col_letter = get_column_letter(col_idx) if get_column_letter else str(col_idx)
                    pf = pf_by_id.get((sh, pf_id)) or pf_by_field.get((sh, lf)) or pf_by_field.get(("", lf))
                    prot_type = entry.protection_type or (pf.protection_type if pf else None)
                    is_num = prot_type == "numeric_scaled"

                    expected_val = entry.get_decoded_masked_value()
                    expected_pool = expected_masked_pool.get((sh, pf_id)) or expected_masked_pool.get((sh, lf), set())

                    if kt not in returned_rows_by_key:
                        continue

                    row_indices = returned_rows_by_key[kt]
                    if len(row_indices) > 1:
                        for r_idx in row_indices:
                            actual_val = ws.cell(row=r_idx, column=col_idx).value
                            change_list.cell_changes.append(
                                CellChange(
                                    sheet=sheet_name,
                                    row_idx=r_idx,
                                    col_idx=col_idx,
                                    col_letter=col_letter,
                                    logical_field=lf,
                                    protected_field_id=entry.protected_field_id,
                                    protection_type=prot_type,
                                    record_key=kt,
                                    before_value=expected_val,
                                    after_value=actual_val,
                                    status=ChangeStatus.RECORD_IDENTITY_CONFLICT,
                                    notes="Duplicate record identity key in returned sheet",
                                )
                            )
                        continue

                    r_idx = row_indices[0]
                    actual_val = ws.cell(row=r_idx, column=col_idx).value

                    # Level 1 Gate: Compare actual_val with expected_val
                    if _values_are_equivalent(expected_val, actual_val, is_numeric=is_num):
                        change_list.cell_changes.append(
                            CellChange(
                                sheet=sheet_name,
                                row_idx=r_idx,
                                col_idx=col_idx,
                                col_letter=col_letter,
                                logical_field=lf,
                                protected_field_id=entry.protected_field_id,
                                protection_type=prot_type,
                                record_key=kt,
                                before_value=expected_val,
                                after_value=actual_val,
                                status=ChangeStatus.UNCHANGED,
                                notes="Expected masked value matches exactly",
                            )
                        )
                    else:
                        actual_str = str(actual_val).strip() if actual_val is not None else ""
                        is_swapped = actual_str in expected_pool

                        notes = "Protected cell value was modified"
                        if is_swapped:
                            notes = "PROTECTED VALUE SWAP: Cell matches a masked value belonging to another record"

                        change_list.cell_changes.append(
                            CellChange(
                                sheet=sheet_name,
                                row_idx=r_idx,
                                col_idx=col_idx,
                                col_letter=col_letter,
                                logical_field=lf,
                                protected_field_id=entry.protected_field_id,
                                protection_type=prot_type,
                                record_key=kt,
                                before_value=expected_val,
                                after_value=actual_val,
                                status=ChangeStatus.PROTECTED_VALUE_CONFLICT,
                                notes=notes,
                            )
                        )

            # Case B: No Record Identity Available
            else:
                cols_to_check = []
                if change_list.matched_columns_by_pf_id.get(sheet_name):
                    for pf_key, col_idx in change_list.matched_columns_by_pf_id[sheet_name].items():
                        pf_obj = pf_by_id.get((sheet_name, pf_key)) or pf_by_field.get((sheet_name, pf_key)) or pf_by_field.get(("", pf_key))
                        lf_name = pf_obj.logical_field if pf_obj else pf_key
                        cols_to_check.append((pf_key, lf_name, col_idx, pf_obj))
                else:
                    for lf, col_idx in change_list.matched_columns[sheet_name].items():
                        pf_obj = pf_by_field.get((sheet_name, lf)) or pf_by_field.get(("", lf))
                        cols_to_check.append((lf, lf, col_idx, pf_obj))

                for pf_key, lf, col_idx, pf in cols_to_check:
                    col_letter = get_column_letter(col_idx) if get_column_letter else str(col_idx)
                    entries = ledger_entries_by_sheet_pf_id.get((sheet_name, pf_key))
                    if not entries:
                        entries = ledger_entries_by_sheet_field.get((sheet_name, lf), [])
                    pf = pf_by_field.get((sheet_name, lf)) or pf_by_field.get(("", lf))
                    prot_type = pf.protection_type if pf else None
                    is_num = prot_type == "numeric_scaled"

                    start_r = returned_scope.data_start_row
                    end_r = returned_scope.data_end_row or (start_r - 1)

                    for offset, r_idx in enumerate(range(start_r, end_r + 1)):
                        actual_val = ws.cell(row=r_idx, column=col_idx).value
                        if offset < len(entries):
                            expected_val = entries[offset].get_decoded_masked_value()
                            entry_pf_id = entries[offset].protected_field_id
                        else:
                            expected_val = None
                            entry_pf_id = None

                        if _values_are_equivalent(expected_val, actual_val, is_numeric=is_num):
                            status = ChangeStatus.UNCHANGED
                            notes = "Positional match (no record identity)"
                        else:
                            status = ChangeStatus.REQUIRE_CONFIRMATION
                            notes = "Value modified without reliable Record Identity; requires human confirmation"

                        change_list.cell_changes.append(
                            CellChange(
                                sheet=sheet_name,
                                row_idx=r_idx,
                                col_idx=col_idx,
                                col_letter=col_letter,
                                logical_field=lf,
                                protected_field_id=entry_pf_id,
                                protection_type=prot_type,
                                record_key=f"row_{r_idx}",
                                before_value=expected_val,
                                after_value=actual_val,
                                status=status,
                                notes=notes,
                            )
                        )

    finally:
        if close_needed and hasattr(wb, "close"):
            wb.close()

    change_list.compute_summary()
    return change_list
