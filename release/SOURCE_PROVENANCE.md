# Independent V1 source snapshot provenance

Snapshot: Excel Privacy Tool V1.0.0, maintained by LunaBella.
Upstream comparison reference: e5b1ff61082ac4dede9c015e0b349c643635114b.
Prior file-level audit: 2026-10-19; publication review: 2026-10-20. No license grant from kwen1510 is assumed.

Included application code:

* privacy_handler/excel_privacy_tool.py: independently rewritten boundary,
  package inspection, metadata sanitation/audit, output and integrity utilities.
  The post-upstream locally developed mask_workbook_v1 orchestration is retained,
  with checked exports and directory isolation. None of the 49 upstream
  top-level function/class bodies are retained unchanged. Legacy upstream CLI
  commands are absent.
* privacy_app.py, gui_controller.py, identity.py, key_detector.py,
  numeric_masking.py, conflict_engine.py, restore_engine.py, prompt_generator.py,
  recovery_ledger.py, table_scope.py, packaging_smoke.py: local post-upstream
  implementations identified in the preceding audit; targeted changes are
  documented by the isolated worktree diff and release changelog.
* privacy_handler/README.md, project README, dependency lists and release docs
  are newly written for V1. approved_checksums.json is freshly generated from
  the exact included resource list. The legacy JSON schema is not shipped.
* tests/test_preflight.py has been independently rewritten; its three copied
  upstream fixture helpers are absent. The retained post-upstream tests have
  no whole function/class AST matches to the upstream tests. test_excel_privacy_tool.py
  is historical CLI material and is excluded from the snapshot.

Excluded, preserved locally/history: configs/, skills/ (including all 16 template
files), historical docs, archive/, examples/, originals, ledgers, keys, restored
outputs, development caches, build outputs and the original Git history.
A fresh Git repository may contain only the explicit staged allowlist. Merely
starting a new repository is not used as the reason for claiming independent code.

The staging tool compares Python ASTs against the locally retained upstream Git
objects and emits a file hash manifest. This is an engineering provenance check,
not proof of all authors' title or a legal clean-room certification. The MIT license applies only to the independently maintained contributions
listed in the reviewed manifest. Third-party packages retain their own licenses
and notices. This engineering review does not establish every author's title
or substitute for any required rights-holder authorization.

2026-10-31 final remediation: removed the unused load_legacy_key_csv reader,
RecoveryMaterial container and load_recovery_material dispatch/fallback, plus
three CSV-only compatibility tests. GUI, Compare and Restore already use
load_encrypted_ledger directly. Non-.enc inputs are now rejected before key
access; decryption failures never fall back to CSV. The V1 encrypted protocol,
entry identity snapshots, conflict decisions and restore implementation are
unchanged. The previously flagged 60-token CSV sequence is not distributed.
The final source manifest records the resulting snapshot. Third-party notices
and necessary upstream attribution remain intact. No upstream license grant
is asserted; static comparison is not a legal clean-room certification.
