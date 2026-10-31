# Excel Privacy Tool V1.0.0

Maintainer: LunaBella. Updated 2026-10-31.

- Reimplemented upstream-origin workbook boundary utilities; retained the local
  V1 pipeline and safety modules. Legacy CLI/Skill scaffolding is not shipped.
- Metadata cleanup errors abort export. Serialized core/custom metadata is
  checked before the output workbook is created.
- V1 consistently rejects .xlsm and VBA payloads.
- Restore rechecks current file contents and uses the same in-memory snapshot
  for identity, comparison and point restore; stale preview plans are advisory.
- Create formula recalculation settings when absent. Actual Excel recalculation
  was separately confirmed in the user-reported Windows Excel 2016 sample acceptance.
- Preserve the existing persisted-workbook identity snapshot fix. Historical
  diagnostics confirm inconsistent old snapshots, not AI ID edits.
- Add source allowlist, third-party notices and native platform acceptance tools.

Publication remains pending. See ACCEPTANCE.md for verified gate results.

- Record completed Windows Excel 2016 sample acceptance and LunaBella branding.
- Replace the macOS bundle identifier with io.lunabella.excelprivacy.

- V1 仅接受 `.enc` 加密恢复账本；已移除旧 CSV 恢复兼容和解密失败后的格式回退。既有 V1 加密账本继续使用原协议。

- Pin Windows packaging Python to the previously accepted 3.12.10 runtime.
