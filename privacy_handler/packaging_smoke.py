"""Internal-only packaged-app round-trip gate using the synthetic HR fixture."""

from __future__ import annotations

import json
from pathlib import Path

from privacy_handler.gui_controller import GUIWorkflowController


HEADER_ROWS = {"人员历史": 2, "人员快照": 1, "薪酬明细": 3}
RECORD_KEYS = {
    "人员历史": ("A", "B"),
    "人员快照": ("A",),
    "薪酬明细": ("A", "B"),
}


def _write_report(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _controller(data_root: Path) -> GUIWorkflowController:
    return GUIWorkflowController(
        safe_dir=data_root / "safe_for_codex",
        restored_dir=data_root / "restored_outputs",
        private_keys_dir=data_root / "private_keys",
    )


def run_ui_phase(report_path: Path) -> None:
    """Initialize the actual packaged Tcl/Tk widgets, not just headless imports."""
    from privacy_app import PrivacyApp

    app = PrivacyApp()
    try:
        app.update()
        if not app.winfo_exists() or not app.winfo_children():
            raise RuntimeError("Packaged UI did not initialize")
        _write_report(report_path, {
            "native_tk_initialized": True,
            "tcl_version": str(app.tk.call("info", "patchlevel")),
            "tk_version": str(app.tk.call("package", "provide", "Tk")),
            "windowing_system": str(app.tk.call("tk", "windowingsystem")),
        })
    finally:
        app.destroy()


def run_mask_phase(source_path: Path, data_root: Path, report_path: Path) -> None:
    """Create one fully configured synthetic run inside an external data root."""
    ctl = _controller(data_root.resolve())
    ok, message = ctl.load_source_workbook(source_path.resolve())
    if not ok:
        raise RuntimeError(message)

    for sheet_name, header_row in HEADER_ROWS.items():
        ok, message = ctl.set_sheet_header_row(sheet_name, header_row)
        if not ok:
            raise RuntimeError(message)

    # Protect every synthetic identity field, including the later processed key.
    choices = {
        "人员历史": (["B", "C", "V", "W", "X", "Y", "Z"], ["N", "O", "P", "Q", "R", "S", "T"]),
        "人员快照": (["A", "B", "V", "W"], ["I", "J", "K", "N", "O", "P"]),
        "薪酬明细": (["B", "C", "M"], ["G", "I"]),
    }
    for name, (text_columns, numeric_columns) in choices.items():
        for letter in text_columns:
            ctl.set_column_protection(name, letter, "text")
        for letter in numeric_columns:
            ctl.set_column_protection(name, letter, "numeric")
    ctl.select_sheet("隐藏配置", False)
    ctl.select_sheet("原有汇总", False)
    ctl.set_column_local_group("薪酬明细", "G", "A")
    ctl.set_column_local_group("薪酬明细", "I", "B")
    ctl.assign_shared_group("共享组-内部打包验证", [("人员历史", "O"), ("人员快照", "J")])

    for sheet_name, columns in RECORD_KEYS.items():
        for column in columns:
            ctl.toggle_column_record_key(sheet_name, column, True)
        ok, message = ctl.confirm_sheet_configuration(sheet_name)
        if not ok:
            raise RuntimeError(message)

    result = ctl.execute_mask()
    ledger = result["ledger"]
    protected_ids = {
        field.protected_field_id
        for field in ledger.protected_fields
        if field.protected_field_id
    }
    duplicate_amount_ids = sorted(
        field_id for field_id in protected_ids if field_id in {"薪酬明细:G", "薪酬明细:I"}
    )
    if duplicate_amount_ids != ["薪酬明细:G", "薪酬明细:I"]:
        raise RuntimeError("Duplicate-header physical identities were not preserved")

    _write_report(
        report_path,
        {
            "phase": "mask",
            "masked_workbook_path": str(Path(result["masked_workbook_path"]).resolve()),
            "ledger_entries": len(ledger.entries),
            "duplicate_amount_ids": duplicate_amount_ids,
            "data_root": str(data_root.resolve()),
        },
    )


def run_resume_phase(
    returned_path: Path,
    foreign_path: Path,
    data_root: Path,
    report_path: Path,
) -> None:
    """In a fresh process, resolve the prior Run, compare, restore, and reject a foreign file."""
    ctl = _controller(data_root.resolve())
    ok, message = ctl.load_returned_workbook(returned_path.resolve())
    if not ok:
        raise RuntimeError(message)
    compare = ctl.execute_compare()
    restore = ctl.execute_restore()

    foreign_ctl = _controller(data_root.resolve())
    foreign_ok, _ = foreign_ctl.load_returned_workbook(foreign_path.resolve())
    if foreign_ok:
        raise RuntimeError("Foreign workbook unexpectedly matched a local Run")

    if compare.can_safe_restore_count != compare.total_protected_cells_count:
        raise RuntimeError("Clean packaged round-trip was not fully restorable")
    if restore.restored_cells_count != restore.total_protected_cells_count:
        raise RuntimeError("Packaged restore did not restore every protected cell")

    _write_report(
        report_path,
        {
            "phase": "resume",
            "matched_existing_run": True,
            "safe_restore_count": compare.can_safe_restore_count,
            "total_protected_cells": compare.total_protected_cells_count,
            "restored_cells": restore.restored_cells_count,
            "protected_conflicts": compare.protected_conflict_count,
            "requires_confirmation": compare.require_confirmation_count,
            "foreign_run_blocked": not foreign_ok,
            "restored_workbook_path": str(restore.restored_workbook_path.resolve()),
        },
    )
