# V1.0 independent release contract

Version: V1.0.0; source audit 2026-10-31.
This is an implementation contract, not a claim of clean-room provenance or release approval.

## Scope and preservation

Keep the existing Python/Tkinter application, V1 mask pipeline, token allocator,
numeric masking, encrypted recovery ledger, Compare and point Restore modules.
Replace the upstream-origin utility/CLI implementation in excel_privacy_tool.py.
Retain mask_workbook_v1, which was added locally after the upstream baseline,
with targeted safety corrections. Its public arguments and result keys remain.
The upstream legacy mask/unmask/anonymise/dedupe CLI and scaffold installer are
outside the V1 desktop contract and excluded from the independent distribution.
The original worktree and all Git history remain preserved.

## Interfaces

* GUI: PrivacyToolError, assert_supported_workbook, detect_unsupported_features,
  mask_workbook_v1, verify_self_before_sensitive.
* Restore: load_workbook_private, resolve_path, check_not_overwrite_source,
  sanitize_workbook_metadata and safe workbook serialization.
* Integrity: LOCKED_FILES, approved_checksums.json; verify exact file hashes before
  sensitive GUI operations. No checksum auto-approval in the shipped application.
* The development command supports only verify-self. Historical CLI tests remain
  in the original history/worktree and are not release acceptance tests.

## Invariants

1. V1 accepts .xlsx only. Reject .xlsm and VBA payloads even under an .xlsx name.
2. Block PivotTable/cache, slicers, connections, custom XML and unsupported payloads.
3. Clear identifying core/custom properties. An exception or incomplete clear
   aborts export. Audit serialized metadata before committing an output file.
4. Field tokens, full Run ID and encrypted ledger schema remain compatible.
   FP8 is only an auxiliary clue; ambiguous batches stop automatic recovery.
5. Retain per-sheet physical column identity, composite record keys, explicit
   numeric groups and the post-serialization record-key snapshot fix.
6. Restore only cells whose current identity and masked values are proven safe;
   never guess rows or reverse numeric scales. Preserve AI changes outside those
   cells. Report all conflict/unconfirmed coverage.
7. Never overwrite an original/returned workbook. Do not place keys in masked
   output directories. Never export plaintext ledgers or secrets.
8. Request Excel recalculation on load; Python does not calculate formulas.
9. No claim of universal workbook sanitization or advanced-object fidelity.

## Release gates

Independent-source allowlist and dependency notices; synthetic package-residue
tests; historical identity-snapshot diagnostics; full applicable regression
suite; macOS arm64 frozen fresh-process roundtrip; Windows x64 native build,
Tk/DLL/Chinese path/fresh-process roundtrip; actual Excel 2016 open/recalculate
and manual verification. Simulations do not satisfy native/Excel gates.
Public publication remains pending the release gates and maintainer approval.
