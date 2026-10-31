"""GUI Workflow Controller and state management for Excel Privacy Tool.

This module provides a headless controller that encapsulates all business logic,
state management, validation, ledger resolution, background threading, and production
API orchestration for privacy_app.py.

Core Principles:
1. Production API Integration: Directly drives mask_workbook_v1, compare_workbook_against_ledger,
   and restore_workbook_v1.
2. Diagnostic Privacy: Never exposes secret keys, scale factors k, or plaintext record identities.
3. Human-Readable Error Messaging: Converts technical engine exceptions into polite, clear Chinese instructions.
4. Testability: All states, validations, and workflows can be thoroughly verified headless without Tkinter widgets.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import io
import logging
import os
from pathlib import Path
import re
import sys
import threading
from typing import Any, Callable

try:
    import openpyxl
except Exception:  # pragma: no cover
    openpyxl = None

from privacy_handler.conflict_engine import (
    ChangeList,
    ChangeStatus,
    ConflictEngineError,
    compare_workbook_against_ledger,
)
from privacy_handler.excel_privacy_tool import (
    PrivacyToolError,
    assert_supported_workbook,
    detect_unsupported_features,
    mask_workbook_v1,
    verify_self_before_sensitive,
)
from privacy_handler.identity import (
    IdentityError,
    compute_run_fingerprint,
    read_workbook_identity,
    resolve_ledger_by_fingerprint,
    TOKEN_SEARCH_REGEX,
)
from privacy_handler.key_detector import analyze_sheet_key_candidates, is_empty_value
from privacy_handler.recovery_ledger import (
    RecoveryLedger,
    RecoveryLedgerError,
    load_encrypted_ledger,
)
from privacy_handler.restore_engine import (
    RestoreEngineError,
    RestoreResult,
    analyze_restore_coverage,
    restore_workbook_v1,
)
from privacy_handler.table_scope import (
    ColumnDescriptor,
    SheetProtectionScope,
    TableScope,
    extract_table_scope,
    validate_composite_key,
)

logger = logging.getLogger(__name__)


# Development mode keeps the existing repository-relative behavior. Frozen
# applications store Ledgers and outputs outside their installation directory.
APP_DIR = Path(__file__).resolve().parents[1]


def default_app_data_root(
    *,
    frozen: bool | None = None,
    platform_name: str | None = None,
    home: Path | str | None = None,
    local_app_data: Path | str | None = None,
) -> Path:
    """Return the writable data root without migrating any legacy files."""
    is_frozen = bool(getattr(sys, "frozen", False)) if frozen is None else frozen
    platform_value = sys.platform if platform_name is None else platform_name
    if is_frozen and platform_value == "darwin":
        home_path = Path.home() if home is None else Path(home)
        return (home_path / "Library" / "Application Support" / "ExcelPrivacyTool").resolve()
    if is_frozen and platform_value == "win32":
        forbidden_roots = (
            Path(sys.executable).resolve().parent,
            Path(getattr(sys, "_MEIPASS", APP_DIR)).resolve(),
            APP_DIR.resolve(),
        )

        def inside_application(path: Path) -> bool:
            return any(path.is_relative_to(root) for root in forbidden_roots)

        configured_root = local_app_data if local_app_data is not None else os.environ.get("LOCALAPPDATA")
        if configured_root and Path(configured_root).is_absolute():
            candidate = (Path(configured_root) / "ExcelPrivacyTool").resolve()
            if not inside_application(candidate):
                return candidate
        home_path = Path.home() if home is None else Path(home)
        fallback = (home_path / "AppData" / "Local" / "ExcelPrivacyTool").resolve()
        if inside_application(fallback):
            raise RuntimeError("No safe per-user application data directory is available")
        return fallback
    return APP_DIR


DEFAULT_DATA_ROOT = default_app_data_root()
DEFAULT_SAFE_DIR = DEFAULT_DATA_ROOT / "safe_for_codex"
DEFAULT_RESTORED_DIR = DEFAULT_DATA_ROOT / "restored_outputs"
DEFAULT_PRIVATE_KEYS_DIR = DEFAULT_DATA_ROOT / "private_keys"


@dataclass
class ColumnConfigState:
    """User configuration state for a single column."""

    column_letter: str
    column_index: int
    header: str
    raw_header: str | None = None
    is_hidden: bool = False
    is_recommended: bool = False
    is_numeric_candidate: bool = False

    # Protection choice: "none", "text", "numeric"
    protection_type: str = "none"

    # Numeric analysis group name (default "可比组 A")
    analysis_group: str = "可比组 A"

    # Numeric analysis group: local group within sheet (default "A")
    local_group: str = "A"
    # Cross-sheet shared group ID if assigned by user (e.g. "shared_1"), else None
    shared_group_id: str | None = None

    # Whether selected as record identity (key) component
    is_record_key: bool = False

    def __post_init__(self):
        val_clean = str(self.analysis_group).strip()
        if val_clean.startswith("共享组"):
            self.shared_group_id = val_clean
        elif "可比组" in val_clean:
            letter = val_clean.replace("可比组", "").strip()
            self.local_group = letter if letter in ["A", "B", "C", "D"] else "A"
            self.shared_group_id = None
        elif val_clean in ["A", "B", "C", "D"]:
            self.local_group = val_clean
            self.shared_group_id = None
            self.analysis_group = f"可比组 {val_clean}"
        elif val_clean:
            self.shared_group_id = val_clean

    @property
    def display_name(self) -> str:
        h = self.header if self.header else "[空表头]"
        hidden_suffix = " [隐藏列]" if self.is_hidden else ""
        return f"{self.column_letter}｜{h}{hidden_suffix}"


@dataclass
class SheetConfigState:
    """User configuration state for a single worksheet."""

    sheet_name: str
    is_selected: bool = False
    header_row: int = 1
    columns: list[ColumnConfigState] = field(default_factory=list)

    # Key validation feedback
    key_validation_status: str = "尚未选择记录识别列。请使用系统推荐或手动选择，验证通过后才能确认本表。"
    key_validation_valid: bool = False
    key_warning_message: str = ""
    recommended_key_desc: str = ""  # e.g. "推荐：A｜员工工号"
    recommended_key_col: str = ""   # e.g. "A"

    # Confirmation state: user must explicitly confirm sheet configuration before masking
    is_confirmed: bool = False

    @property
    def is_configured(self) -> bool:
        """Alias for is_confirmed for UI badge display."""
        return self.is_confirmed

    def get_column(self, col_letter: str) -> ColumnConfigState | None:
        for c in self.columns:
            if c.column_letter.upper() == col_letter.upper():
                return c
        return None

    def get_record_key_fields(self) -> list[str]:
        return [c.header for c in self.columns if c.is_record_key]

    def get_record_key_columns(self) -> list[str]:
        """Return physical column identities for execution-time matching."""
        return [c.column_letter for c in self.columns if c.is_record_key]

    def get_text_protection_fields(self) -> list[str]:
        return [c.header for c in self.columns if c.protection_type == "text"]

    def get_text_protection_columns(self) -> list[str]:
        """Return physical text-column identities, preserving duplicate headers."""
        return [c.column_letter for c in self.columns if c.protection_type == "text"]

    def get_numeric_protection_fields(self) -> list[str]:
        return [c.header for c in self.columns if c.protection_type == "numeric"]

    def get_numeric_protection_columns(self) -> list[str]:
        """Return physical numeric-column identities, preserving duplicate headers."""
        return [c.column_letter for c in self.columns if c.protection_type == "numeric"]


@dataclass
class CompareSummary:
    """Human-readable summary of AI returned file changes."""

    can_safe_restore_count: int = 0
    protected_conflict_count: int = 0
    require_confirmation_count: int = 0
    added_rows_count: int = 0
    deleted_rows_count: int = 0
    added_columns_count: int = 0
    deleted_columns_count: int = 0
    has_conflicts: bool = False
    conflict_details: list[dict[str, Any]] = field(default_factory=list)

    # UX Closeout additions: Reason breakdown summary & top reason description
    reason_breakdown: dict[str, int] = field(default_factory=dict)
    primary_unconfirmed_reason: str = ""
    grouped_conflict_rows: list[dict[str, Any]] = field(default_factory=list)
    total_protected_cells_count: int = 0
    unresolved_by_sheet: dict[str, int] = field(default_factory=dict)


def translate_security_error(exc: Exception) -> str:
    """Convert raw internal error into user-friendly polite Chinese message."""
    msg = str(exc)
    if "Unsupported workbook features detected" in msg or "PivotTable" in msg or "Slicer" in msg:
        return "这个 Excel 包含当前版本暂不支持的高级对象（如数据透视表、切片器或外部连接）。为了避免文件损坏，本次操作已停止，原文件没有被修改。"
    if "RUN_IDENTITY_MISMATCH" in msg or "RUN_IDENTITY_UNVERIFIED" in msg or "Cannot safely restore" in msg:
        return "无法确认这个文件属于当前脱敏任务，因此没有执行恢复。请确认选择的是本次脱敏后由 AI 处理返回的对应文件。"
    if "No recovery ledger found" in msg or "Ambiguous FP8" in msg:
        return "本地未找到该文件对应的隐私恢复账本，或存在多个任务歧义，无法自动恢复。"
    if "PROTECTED_VALUE_CONFLICT" in msg:
        return "AI 修改了受保护字段，为保护原始业务数据，系统不会静默覆盖。"
    return f"操作未完成：{msg}"


class GUIWorkflowController:
    """Headless state controller for the Excel Privacy Tool GUI."""

    def __init__(
        self,
        safe_dir: Path | str | None = None,
        restored_dir: Path | str | None = None,
        private_keys_dir: Path | str | None = None,
    ):
        self.safe_dir = Path(safe_dir).resolve() if safe_dir else DEFAULT_SAFE_DIR
        self.restored_dir = Path(restored_dir).resolve() if restored_dir else DEFAULT_RESTORED_DIR
        self.private_keys_dir = Path(private_keys_dir).resolve() if private_keys_dir else DEFAULT_PRIVATE_KEYS_DIR

        # Step 1: Source Workbook State
        self.source_path: Path | None = None
        self.sheet_configs: dict[str, SheetConfigState] = {}
        self.available_sheets: list[str] = []

        # Step 5 & 6: Mask Result State
        self.last_mask_result: dict[str, Any] | None = None
        self.ai_prompt: str = ""
        self.last_masked_path: Path | None = None

        # Step 8-12: AI Returned & Restore State
        self.returned_path: Path | None = None
        self.resolved_run_id: str | None = None
        self.resolved_batch_tag: str | None = None
        self.resolved_ledger_path: Path | None = None
        self.resolved_secret_key: bytes | None = None
        self.last_change_list: ChangeList | None = None
        self.last_compare_summary: CompareSummary | None = None
        self.last_restore_result: RestoreResult | None = None

        # Lock / busy state for UI safety
        self.is_busy: bool = False
        self.status_message: str = "请选择需要脱敏的 Excel 文件"

    # --------------------------------------------------------------------------
    # Step 1 & 2: Load Source Workbook
    # --------------------------------------------------------------------------

    def load_source_workbook(self, file_path: Path | str) -> tuple[bool, str]:
        """Inspect source workbook, run pre-flight gate, and initialize sheet configurations."""
        path = Path(file_path).resolve()
        if not path.exists():
            return False, f"文件不存在：{path.name}"

        # 1. Pre-flight Unsupported parts gate
        try:
            assert_supported_workbook(path, action_name="脱敏")
        except Exception as exc:
            return False, translate_security_error(exc)

        # 2. Inspect sheets
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
            sheetnames = list(wb.sheetnames)
            wb.close()
        except Exception as exc:
            return False, f"读取 Excel 文件失败：{exc}"

        if not sheetnames:
            return False, "Excel 文件中没有找到有效的工作表"

        self.source_path = path
        self.available_sheets = sheetnames
        self.sheet_configs.clear()

        # Build initial sheet configurations
        for idx, sname in enumerate(sheetnames):
            # Recommend: single sheet is auto-selected; multi-sheet default all selected
            cfg = SheetConfigState(sheet_name=sname, is_selected=True, header_row=1)
            self.sheet_configs[sname] = cfg
            self.refresh_sheet_columns(sname, header_row=1)

        self.status_message = f"工作簿已载入：{path.name}（共 {len(sheetnames)} 个工作表）"
        return True, self.status_message

    def select_sheet(self, sheet_name: str, selected: bool) -> None:
        if sheet_name in self.sheet_configs:
            self.sheet_configs[sheet_name].is_selected = selected

    def set_sheet_header_row(self, sheet_name: str, header_row: int) -> tuple[bool, str]:
        """Update header row for a sheet and re-scan columns immediately."""
        if sheet_name not in self.sheet_configs:
            return False, f"未找到工作表：{sheet_name}"
        if header_row < 1:
            return False, "表头行必须大于等于 1"

        cfg = self.sheet_configs[sheet_name]
        cfg.header_row = header_row
        cfg.is_confirmed = False  # Invalidate confirmation upon header row change
        return self.refresh_sheet_columns(sheet_name, header_row=header_row)

    def refresh_sheet_columns(self, sheet_name: str, header_row: int) -> tuple[bool, str]:
        """Extract columns from sheet using TableScope and auto-populate recommendations."""
        if not self.source_path or not self.source_path.exists():
            return False, "源文件未载入"

        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return False, f"未找到工作表配置：{sheet_name}"

        # Preserve previous column choices if they existed (key invariant)
        prev_cols_by_key: dict[str, ColumnConfigState] = {}
        for c in cfg.columns:
            prev_cols_by_key[f"{c.column_letter}:{c.header}"] = c
            prev_cols_by_key[c.column_letter] = c

        try:
            wb = openpyxl.load_workbook(self.source_path, data_only=False)
            if sheet_name not in wb.sheetnames:
                wb.close()
                return False, f"工作表中不存在 {sheet_name}"
            ws = wb[sheet_name]
            scope = extract_table_scope(ws, header_row=header_row)

            # Analyze key candidate statistics
            key_stats = analyze_sheet_key_candidates(ws, header_row=header_row)
            wb.close()
        except Exception as exc:
            return False, f"读取表头列清单失败：{exc}"

        new_columns: list[ColumnConfigState] = []
        strong_candidates: list[tuple[str, str]] = []  # list of (letter, header)
        for col_desc in scope.selected_columns:
            # Check if this column had existing user configuration matching both letter and header
            prev = prev_cols_by_key.get(f"{col_desc.column_letter}:{col_desc.header}")

            if prev is not None:
                prot_type = prev.protection_type
                grp = prev.analysis_group
                is_key = prev.is_record_key
            else:
                prot_type = "none"
                if col_desc.recommended_protection:
                    prot_type = "text"
                elif col_desc.is_numeric_candidate:
                    prot_type = "none"
                grp = "可比组 A"
                is_key = False

            # Key candidate recommendation
            k_stat = key_stats.get(col_desc.header)
            is_strong_key = k_stat.is_strong_candidate if k_stat else False
            if is_strong_key:
                strong_candidates.append((col_desc.column_letter, col_desc.header))

            col_state = ColumnConfigState(
                column_letter=col_desc.column_letter,
                column_index=col_desc.column_index,
                header=col_desc.header,
                raw_header=col_desc.raw_header,
                is_hidden=col_desc.is_hidden,
                is_recommended=col_desc.recommended_protection,
                is_numeric_candidate=col_desc.is_numeric_candidate,
                protection_type=prot_type,
                analysis_group=grp,
                is_record_key=is_key,
            )
            new_columns.append(col_state)

        cfg.columns = new_columns
        if strong_candidates:
            rec_letter, rec_header = strong_candidates[0]
            cfg.recommended_key_desc = f"{rec_letter}｜{rec_header}"
            cfg.recommended_key_col = rec_letter
        else:
            cfg.recommended_key_desc = ""
            cfg.recommended_key_col = ""

        self.validate_sheet_record_identity(sheet_name)
        return True, f"成功刷新 {sheet_name} 表头列（第 {header_row} 行，共 {len(new_columns)} 列）"

    # --------------------------------------------------------------------------
    # Step 3: Column Protection & Record Key Selection
    # --------------------------------------------------------------------------

    def confirm_sheet_configuration(self, sheet_name: str) -> tuple[bool, str]:
        """Validate and explicitly confirm configuration for a sheet."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return False, f"未找到工作表：{sheet_name}"
        cfg.is_confirmed = False
        if cfg.header_row < 1:
            return False, f"工作表【{sheet_name}】的表头行设置无效"
        if not cfg.columns:
            return False, f"工作表【{sheet_name}】没有可配置的数据字段"
        t_cols = cfg.get_text_protection_fields()
        n_cols = cfg.get_numeric_protection_fields()
        if not t_cols and not n_cols:
            return False, f"请在工作表【{sheet_name}】中至少选择一个保护字段（文本保护或数值保护）"
        identity_result = self.validate_sheet_record_identity(sheet_name)
        if not identity_result.get("is_valid"):
            return False, cfg.key_validation_status
        # Check numeric protection eligibility: non-numeric columns cannot have numeric protection
        for col in cfg.columns:
            if col.protection_type == "numeric" and not col.is_numeric_candidate:
                return False, f"列【{col.display_name}】包含非数值数据，不可设置为数值保护"
        cfg.is_confirmed = True
        return True, f"✓ 工作表【{sheet_name}】配置已确认"

    def apply_recommended_key(self, sheet_name: str) -> bool:
        """Apply the recommended record identification column upon explicit user action."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg or not cfg.recommended_key_col:
            return False
        # Set only the recommended column as record key
        for c in cfg.columns:
            c.is_record_key = (c.column_letter.upper() == cfg.recommended_key_col.upper())
        cfg.is_confirmed = False  # Invalidate confirmation upon modification
        self.validate_sheet_record_identity(sheet_name)
        return True

    def set_column_protection(
        self, sheet_name: str, col_letter: str, protection_type: str, analysis_group: str = "可比组 A"
    ) -> None:
        """Set protection type ('none', 'text', 'numeric') and optional analysis group."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return
        col = cfg.get_column(col_letter)
        if col:
            # Check eligibility: only numeric candidates can select numeric
            if protection_type == "numeric" and not col.is_numeric_candidate:
                protection_type = "none"
            col.protection_type = protection_type
            col.analysis_group = analysis_group.strip() or "可比组 A"
            col.__post_init__()
            cfg.is_confirmed = False  # Invalidate confirmation upon edit

    def toggle_column_record_key(self, sheet_name: str, col_letter: str, is_key: bool) -> None:
        """Toggle record key component for a column and re-validate immediately."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return
        col = cfg.get_column(col_letter)
        if col:
            col.is_record_key = is_key
            cfg.is_confirmed = False  # Invalidate confirmation upon edit
            self.validate_sheet_record_identity(sheet_name)

    def validate_sheet_record_identity(self, sheet_name: str) -> dict[str, Any]:
        """Validate current selected record identity columns for the sheet."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return {"is_valid": False, "reason": "工作表不存在"}

        selected_keys = cfg.get_record_key_columns()
        if not selected_keys:
            cfg.key_validation_valid = False
            cfg.is_confirmed = False
            cfg.key_validation_status = "尚未选择记录识别列。请使用系统推荐或手动选择，验证通过后才能确认本表。"
            cfg.key_warning_message = cfg.key_validation_status
            return {"is_valid": False, "reason": "no_key", "warning": cfg.key_warning_message}

        if not self.source_path or not self.source_path.exists():
            cfg.key_validation_valid = False
            cfg.is_confirmed = False
            cfg.key_validation_status = "无法完成记录识别验证：源文件不存在。请重新选择文件后再试。"
            return {"is_valid": False, "reason": "source_missing"}

        try:
            wb = openpyxl.load_workbook(self.source_path, data_only=False)
            ws = wb[sheet_name]
            from privacy_handler.table_scope import iter_table_scope_rows
            scope = extract_table_scope(ws, header_row=cfg.header_row)
            rows = iter_table_scope_rows(ws, scope)
            wb.close()

            # TableScope excludes trailing formatted rows. Also exclude wholly
            # empty internal/formatting rows before key validation so only
            # actual data records participate in uniqueness and null checks.
            physical_columns = [column.column_letter for column in cfg.columns]
            rows = [
                row
                for row in rows
                if any(not is_empty_value(row.get(letter)) for letter in physical_columns)
            ]

            val_res = validate_composite_key(rows, selected_keys)
            if val_res["is_valid"]:
                cfg.key_validation_valid = True
                cfg.key_validation_status = "✓ 记录识别验证通过，可以确认本表配置。"
                cfg.key_warning_message = ""
            else:
                cfg.key_validation_valid = False
                cfg.is_confirmed = False
                if val_res.get("null_count", 0) > 0:
                    cfg.key_validation_status = "当前识别字段包含空值，无法可靠识别全部记录，请调整后重试。"
                elif val_res.get("formula_count", 0) > 0:
                    cfg.key_validation_status = "当前识别字段包含公式，无法作为稳定的记录识别，请重新选择。"
                elif val_res.get("unique_count", 0) < val_res.get("total_count", 0):
                    cfg.key_validation_status = "当前识别字段存在重复记录，请增加能够区分记录的字段或重新选择。"
                else:
                    cfg.key_validation_status = "无法完成记录识别验证，请检查字段和数据后重试。"
                cfg.key_warning_message = cfg.key_validation_status
            return val_res
        except Exception as exc:
            cfg.key_validation_valid = False
            cfg.is_confirmed = False
            cfg.key_validation_status = f"无法完成记录识别验证：{exc}"
            cfg.key_warning_message = cfg.key_validation_status
            return {"is_valid": False, "reason": str(exc)}

    # --------------------------------------------------------------------------
    # Step 5: Execute Masking
    # --------------------------------------------------------------------------

    def validate_mask_preconditions(self) -> tuple[bool, str]:
        """Verify that all inputs are valid before triggering mask_workbook_v1."""
        if not self.source_path or not self.source_path.exists():
            return False, "请先选择需要脱敏的 Excel 原始文件。"

        selected_sheets = [s for s in self.sheet_configs.values() if s.is_selected]
        if not selected_sheets:
            return False, "请至少勾选一个需要处理的工作表。"

        total_protected = 0
        for s in selected_sheets:
            if s.header_row < 1:
                return False, f"工作表【{s.sheet_name}】的表头行设置无效。"
            t_cols = s.get_text_protection_fields()
            n_cols = s.get_numeric_protection_fields()
            total_protected += len(t_cols) + len(n_cols)

            # Re-open the current source and validate the current key selection.
            # Never trust a cached is_confirmed/key_validation_valid state here.
            identity_result = self.validate_sheet_record_identity(s.sheet_name)
            if not identity_result.get("is_valid"):
                s.is_confirmed = False
                return False, f"工作表【{s.sheet_name}】：{s.key_validation_status}"

        # Multi-sheet confirmation gate: validation success does not replace
        # the user's explicit per-sheet confirmation.
        unconfirmed = [s.sheet_name for s in selected_sheets if not s.is_confirmed]
        if unconfirmed:
            if len(unconfirmed) == 1:
                return False, f"还有 1 个工作表尚未确认配置：\n{unconfirmed[0]}\n\n请在右侧检查字段后点击【确认本表配置】。"
            names = "、".join(unconfirmed)
            return False, f"还有 {len(unconfirmed)} 个工作表尚未确认配置：\n{names}\n\n请逐个进入工作表检查字段并点击【确认本表配置】。"

        if total_protected == 0:
            return False, "请至少在一个工作表中选择至少一列保护字段（文本保护或金额保护）。"

        try:
            self.safe_dir.mkdir(parents=True, exist_ok=True)
            test_file = self.safe_dir / ".write_test"
            test_file.touch()
            test_file.unlink()
        except Exception as exc:
            return False, f"脱敏保存目录不可写：{self.safe_dir} ({exc})"

        return True, "验证通过"

    def set_column_local_group(self, sheet_name: str, col_letter: str, local_group: str) -> None:
        """Set local group ('A', 'B', 'C', 'D') for a numeric column."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return
        col = cfg.get_column(col_letter)
        if col and col.protection_type == "numeric":
            col.local_group = local_group.strip().upper() if local_group.strip().upper() in ["A", "B", "C", "D"] else "A"
            col.shared_group_id = None
            cfg.is_confirmed = False

    def assign_shared_group(self, group_id: str, members: list[tuple[str, str]]) -> None:
        """Assign multiple columns across sheets to a shared group.

        members: list of (sheet_name, col_letter).
        Invalidates confirmation on all affected sheets.
        """
        affected_sheets: set[str] = set()
        # Find sheets previously in this shared group
        for sname, scfg in self.sheet_configs.items():
            for c in scfg.columns:
                if c.shared_group_id == group_id:
                    affected_sheets.add(sname)
                    c.shared_group_id = None

        # Assign new members
        for sname, col_letter in members:
            scfg = self.sheet_configs.get(sname)
            if scfg:
                c = scfg.get_column(col_letter)
                if c and c.protection_type == "numeric":
                    c.shared_group_id = group_id
                    affected_sheets.add(sname)

        for sname in affected_sheets:
            if sname in self.sheet_configs:
                self.sheet_configs[sname].is_confirmed = False

    def remove_from_shared_group(self, sheet_name: str, col_letter: str) -> None:
        """Remove a column from its shared group, restoring its existing local group."""
        cfg = self.sheet_configs.get(sheet_name)
        if not cfg:
            return
        col = cfg.get_column(col_letter)
        if col and col.shared_group_id:
            col.shared_group_id = None
            cfg.is_confirmed = False

    def get_shared_groups(self) -> dict[str, list[tuple[str, str, str]]]:
        """Return dict of shared_group_id -> list of (sheet_name, col_letter, header)."""
        res: dict[str, list[tuple[str, str, str]]] = {}
        for sname, scfg in self.sheet_configs.items():
            for c in scfg.columns:
                if c.protection_type == "numeric" and c.shared_group_id:
                    res.setdefault(c.shared_group_id, []).append((sname, c.column_letter, c.header))
        return res

    def build_sheet_scopes(self) -> list[SheetProtectionScope]:
        """Build SheetProtectionScope list for mask_workbook_v1."""
        scopes: list[SheetProtectionScope] = []
        for s in self.sheet_configs.values():
            if not s.is_selected:
                continue

            # Execution uses physical column letters. Header text is presentation
            # metadata only and is not unique when a sheet contains duplicate headers.
            t_cols = s.get_text_protection_columns()
            n_cols = s.get_numeric_protection_columns()
            keys = s.get_record_key_columns()

            # Map each numeric column to its effective group identity:
            # - If shared_group_id is set: use shared_group_id (e.g. "shared_1" or "共享组A")
            # - Else: "local::{sheet_name}::{local_group}"
            col_groups: dict[str, str] = {}
            for col in s.columns:
                if col.protection_type == "numeric":
                    if col.shared_group_id:
                        col_groups[col.column_letter] = col.shared_group_id
                    elif col.local_group:
                        col_groups[col.column_letter] = f"local::{s.sheet_name}::{col.local_group}"

            # Legacy fallback group_name for single-group callers
            primary_group = next(iter(col_groups.values()), None) if col_groups else None

            scope = SheetProtectionScope(
                sheet_name=s.sheet_name,
                header_row=s.header_row,
                text_columns=t_cols,
                numeric_columns=n_cols,
                record_key=keys if keys else None,
                numeric_analysis_group=primary_group,
                numeric_column_groups=col_groups,
            )
            scopes.append(scope)
        return scopes

    def execute_mask(self) -> dict[str, Any]:
        """Execute production masking via mask_workbook_v1."""
        verify_self_before_sensitive()
        valid, msg = self.validate_mask_preconditions()
        if not valid:
            raise PrivacyToolError(msg)

        scopes = self.build_sheet_scopes()
        result = mask_workbook_v1(
            input_path=self.source_path,
            output_dir=self.safe_dir,
            private_key_dir=self.private_keys_dir,
            sheet_scopes=scopes,
        )

        self.last_mask_result = result
        self.ai_prompt = result.get("ai_prompt", "")
        self.last_masked_path = Path(result["masked_workbook_path"]).resolve()
        self.status_message = f"脱敏完成：已生成 {self.last_masked_path.name}"
        return result

    # --------------------------------------------------------------------------
    # Step 8-10: Select AI Returned File & Compare
    # --------------------------------------------------------------------------

    def load_returned_workbook(self, file_path: Path | str) -> tuple[bool, str]:
        """Inspect AI-returned workbook and auto-resolve Recovery Ledger from private_keys."""
        path = Path(file_path).resolve()
        if not path.exists():
            return False, f"返回文件不存在：{path.name}"

        # 1. Pre-flight Unsupported parts gate
        try:
            assert_supported_workbook(path, action_name="还原")
        except Exception as exc:
            return False, translate_security_error(exc)

        self.returned_path = path

        # 2. Auto-resolve Recovery Ledger via Run Identity / FP8 clue
        try:
            wb_ident = read_workbook_identity(path)
            run_id = wb_ident.get("run_id")
            ledger_path: Path | None = None

            if run_id:
                cand = self.private_keys_dir / f"{run_id}.recovery_ledger.enc"
                if cand.exists():
                    ledger_path = cand
                    self.resolved_run_id = run_id

            if not ledger_path:
                # Fallback to FP8 search in cells
                found_fp8s: set[str] = set()
                wb_temp = openpyxl.load_workbook(path, read_only=True, data_only=True)
                for sname in wb_temp.sheetnames:
                    ws_temp = wb_temp[sname]
                    for row in ws_temp.iter_rows(values_only=True):
                        for val in row:
                            if val and isinstance(val, str):
                                m = TOKEN_SEARCH_REGEX.search(val)
                                if m:
                                    found_fp8s.add(m.group(1).upper())
                wb_temp.close()

                if len(found_fp8s) == 1:
                    single_fp8 = list(found_fp8s)[0]
                    ledger_path = resolve_ledger_by_fingerprint(self.private_keys_dir, single_fp8)
                    self.resolved_run_id = ledger_path.name[:-len(".recovery_ledger.enc")]

            if not ledger_path or not ledger_path.exists():
                self.resolved_run_id = None
                self.resolved_batch_tag = None
                self.status_message = "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。"
                return False, self.status_message

            self.resolved_ledger_path = ledger_path
            secret_key_file = self.private_keys_dir / f"{self.resolved_run_id}.mask_secret.key"
            if not secret_key_file.exists():
                self.resolved_run_id = None
                self.resolved_batch_tag = None
                self.status_message = "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。"
                return False, self.status_message
            self.resolved_secret_key = secret_key_file.read_bytes().strip()

            fp8 = compute_run_fingerprint(self.resolved_run_id)
            self.resolved_batch_tag = fp8
            self.status_message = f"✓ 已自动匹配该文件对应的本地恢复资料\n脱敏批次：{fp8}"
            return True, self.status_message

        except Exception as exc:
            self.resolved_run_id = None
            self.resolved_batch_tag = None
            self.status_message = "⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。"
            return False, self.status_message

    def execute_compare(self) -> CompareSummary:
        """Execute conflict engine compare against auto-resolved Recovery Ledger."""
        if not self.returned_path or not self.resolved_ledger_path or not self.resolved_secret_key:
            raise PrivacyToolError("请先选择有效的 AI 返回文件并匹配恢复账本。")

        try:
            verify_self_before_sensitive()
            change_list = compare_workbook_against_ledger(
                returned_workbook_or_path=self.returned_path,
                ledger_or_path=self.resolved_ledger_path,
                secret_key=self.resolved_secret_key,
            )
            self.last_change_list = change_list

            ledger = load_encrypted_ledger(
                self.resolved_ledger_path, self.resolved_secret_key
            )
            coverage = analyze_restore_coverage(ledger, change_list)

            # These counts are reconciled against every Ledger entry. In
            # particular, rows lost through a changed Record Identifier and
            # no-key ambiguous values must not be advertised as safely restorable.
            safe_restore = coverage.restorable_cells_count
            protected_conflict = coverage.conflict_cells_count
            req_confirm = coverage.unconfirmed_cells_count

            added_rows = sum(1 for s in change_list.structural_changes if s.change_type == ChangeStatus.ADDED_ROW)
            deleted_rows = sum(1 for s in change_list.structural_changes if s.change_type == ChangeStatus.DELETED_ROW)
            added_cols = sum(1 for s in change_list.structural_changes if s.change_type == ChangeStatus.ADDED_COLUMN)
            deleted_cols = sum(1 for s in change_list.structural_changes if s.change_type == ChangeStatus.DELETED_COLUMN)

            # Collect safe ephemeral conflict items & build reason breakdown
            conflict_details: list[dict[str, Any]] = []
            safe_dict = change_list.to_safe_dict()
            reason_breakdown: dict[str, int] = {}

            # Aggregation bucket: (sheet, field, status_reason) -> count
            grouped_counts: dict[tuple[str, str, str], int] = {}

            for c in safe_dict.get("cell_changes", []):
                st = c.get("status")
                sheet_name = c.get("sheet", "")
                field_name = c.get("logical_field", "")
                rec_id = c.get("record_key", "Record-XX")
                row_idx = c.get("row_idx")
                col_letter = c.get("col_letter", "")
                protection_type = c.get("protection_type", "")

                reason_desc = ""
                detail_reason = ""
                if st == ChangeStatus.PROTECTED_VALUE_CONFLICT.value:
                    reason_desc = "AI 修改了受保护数据"
                    detail_reason = (
                        "AI 修改了受保护数值"
                        if protection_type == "numeric_scale"
                        else "AI 修改了受保护文本"
                    )
                elif st == ChangeStatus.REQUIRE_CONFIRMATION.value:
                    reason_desc = "未设置可靠的记录匹配列"
                    detail_reason = reason_desc
                elif st == ChangeStatus.RECORD_IDENTITY_CONFLICT.value:
                    reason_desc = "记录匹配列存在重复"
                    detail_reason = reason_desc

                if reason_desc:
                    reason_breakdown[reason_desc] = reason_breakdown.get(reason_desc, 0) + 1
                    g_key = (sheet_name, field_name, reason_desc)
                    grouped_counts[g_key] = grouped_counts.get(g_key, 0) + 1

                    conflict_details.append({
                        "sheet": sheet_name,
                        "cell": f"{col_letter}{row_idx}" if col_letter and row_idx else "",
                        "field": f"{col_letter}｜{field_name}" if col_letter else field_name,
                        "record": rec_id,
                        "reason": detail_reason,
                        "policy": "保留 AI 返回值，不自动覆盖" if st == ChangeStatus.PROTECTED_VALUE_CONFLICT.value else "需人工确认",
                        "ai_value": c.get("after_value"),
                        "status": f"{detail_reason}（不自动覆盖）" if st == ChangeStatus.PROTECTED_VALUE_CONFLICT.value else f"{detail_reason}（需人工确认）",
                        "raw_reason": reason_desc,
                    })

            # Structural key changes are not Protected Value Conflicts. Show
            # them explicitly so added/deleted key rows cannot disappear behind
            # the protected-cell conflict list.
            for structural in safe_dict.get("structural_changes", []):
                st = structural.get("change_type")
                if st not in (
                    ChangeStatus.ADDED_ROW.value,
                    ChangeStatus.DELETED_ROW.value,
                    ChangeStatus.RECORD_IDENTITY_CONFLICT.value,
                ):
                    continue
                details = structural.get("details") or {}
                row_indices = details.get("row_indices") or details.get("duplicate_row_indices") or []
                cell_display = f"第{row_indices[0]}行" if row_indices else "原记录"
                if st == ChangeStatus.ADDED_ROW.value:
                    reason = "新增记录，或记录识别值已变更（新键）"
                elif st == ChangeStatus.DELETED_ROW.value:
                    reason = "原记录缺失，或记录识别值已变更（原键）"
                else:
                    reason = "记录识别值重复或无法唯一定位"
                conflict_details.append({
                    "sheet": structural.get("sheet", ""),
                    "cell": cell_display,
                    "field": "记录识别",
                    "record": structural.get("identifier", ""),
                    "reason": reason,
                    "policy": "不猜测原记录，需人工核对",
                    "ai_value": "",
                    "status": f"{reason}（需人工核对）",
                    "raw_reason": "记录识别异常",
                })

            # A Ledger-reconciled aggregate makes the impact range visible even
            # when no CellChange exists (for example, a key-changed/deleted row).
            for sheet_name, fields in coverage.unresolved_by_sheet_field.items():
                for field_name, count in sorted(fields.items()):
                    conflict_details.append({
                        "sheet": sheet_name,
                        "cell": "—",
                        "field": field_name,
                        "record": "",
                        "reason": f"无法唯一定位原记录（{count} 项）",
                        "policy": "保留 AI 返回值，需人工核对",
                        "ai_value": "",
                        "status": "无法安全自动恢复",
                        "raw_reason": "无法安全定位的受保护数据",
                    })

            # Use coverage totals rather than detail-row counts; aggregate rows
            # intentionally summarize many affected cells.
            explicit_no_key_count = reason_breakdown.get("未设置可靠的记录匹配列", 0)
            reason_breakdown = {}
            if protected_conflict:
                reason_breakdown["AI 修改了受保护数据"] = protected_conflict
            if explicit_no_key_count:
                reason_breakdown["未设置可靠的记录匹配列"] = min(
                    explicit_no_key_count, req_confirm
                )
            remaining_unconfirmed = max(req_confirm - explicit_no_key_count, 0)
            if remaining_unconfirmed:
                reason_breakdown["无法安全定位的受保护数据"] = remaining_unconfirmed

            # Format grouped conflict rows: e.g. "Staff｜基本工资", "未设置可靠记录匹配列", "180 项"
            grouped_rows: list[dict[str, Any]] = []
            for (s_name, f_name, r_desc), count in grouped_counts.items():
                grouped_rows.append({
                    "sheet": s_name,
                    "field": f_name,
                    "target": f"{s_name}｜{f_name}",
                    "reason": r_desc,
                    "count": count,
                    "display": f"{count} 项",
                })

            primary_unconfirmed = ""
            if reason_breakdown:
                top_reason, top_count = max(reason_breakdown.items(), key=lambda x: x[1])
                primary_unconfirmed = f"当前主要原因：⚠ {top_reason}（{top_count} 项）"

            summary = CompareSummary(
                can_safe_restore_count=safe_restore,
                protected_conflict_count=protected_conflict,
                require_confirmation_count=req_confirm,
                added_rows_count=added_rows,
                deleted_rows_count=deleted_rows,
                added_columns_count=added_cols,
                deleted_columns_count=deleted_cols,
                has_conflicts=change_list.has_conflicts,
                conflict_details=conflict_details,
                reason_breakdown=reason_breakdown,
                primary_unconfirmed_reason=primary_unconfirmed,
                grouped_conflict_rows=grouped_rows,
                total_protected_cells_count=coverage.total_protected_cells_count,
                unresolved_by_sheet=coverage.unresolved_by_sheet,
            )
            self.last_compare_summary = summary
            self.status_message = f"比对完成：【可自动恢复】{safe_restore} 项，【AI 修改了保护数据】{protected_conflict} 项，【需要人工核对】【无法安全自动判断】{req_confirm} 项"
            return summary

        except Exception as exc:
            raise PrivacyToolError(translate_security_error(exc)) from exc

    # --------------------------------------------------------------------------
    # Step 11 & 12: Execute Restore
    # --------------------------------------------------------------------------

    def execute_restore(self) -> RestoreResult:
        """Execute restore_workbook_v1 with allow_conflicts=True (safe auto-restore)."""
        if not self.returned_path or not self.resolved_ledger_path or not self.resolved_secret_key:
            raise PrivacyToolError("请先选择有效的 AI 返回文件并匹配恢复账本。")

        try:
            verify_self_before_sensitive()
            self.restored_dir.mkdir(parents=True, exist_ok=True)
            res = restore_workbook_v1(
                returned_workbook_path=self.returned_path,
                ledger_path=self.resolved_ledger_path,
                secret_key=self.resolved_secret_key,
                output_dir=self.restored_dir,
                change_list=self.last_change_list,
                allow_conflicts=True,
            )
            self.last_restore_result = res
            if res.skipped_conflicts_count or res.skipped_unconfirmed_count:
                self.status_message = (
                    f"部分恢复完成：已恢复 {res.restored_cells_count} / "
                    f"{res.total_protected_cells_count} 项，仍有 "
                    f"{res.skipped_conflicts_count + res.skipped_unconfirmed_count} 项需人工核对"
                )
            else:
                self.status_message = f"恢复完成：已生成 {res.restored_workbook_path.name}（恢复 {res.restored_cells_count} 项）"
            return res
        except Exception as exc:
            raise PrivacyToolError(translate_security_error(exc)) from exc
