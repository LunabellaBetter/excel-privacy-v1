# Excel Privacy V1 desktop implementation

The application uses local field tokens and an encrypted recovery ledger to
protect selected .xlsx columns before AI editing. Compare and Restore require
verified batch identity and conservatively retain unconfirmed values.

V1 rejects .xlsm, encrypted Office files, legacy .xls, PivotTable/cache,
slicers, external connections/links, custom XML, embedded objects and ActiveX.
Unselected fields, comments, images and hidden worksheets require user review.
Field masking is not a guarantee that an entire workbook is anonymous.

The independently rewritten workbook boundary lives in excel_privacy_tool.py.
The locally developed V1 pipeline and safety modules retain their existing
interfaces. The upstream legacy CLI and Skill template are not distributed.
See docs/V1_RELEASE_CONTRACT.md and release/SOURCE_PROVENANCE.md.

Source integrity verification is a consistency check, not a digital signature.
Run: python privacy_app.py --verify-self
