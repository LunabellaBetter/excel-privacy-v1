# Excel Privacy Tool V1.0.0 acceptance

Maintainer: LunaBella. Evidence reviewed: 2026-10-31.

## Historical native automation (2026-10-19 / 2026-10-20)

- macOS arm64: 304 core regressions and 144 GUI tests in isolated processes passed.
  Repeated Tk 9.1 teardown crashed a single-process run; no full single-process pass is claimed.
- Windows x64: 304 core and 141 isolated GUI tests passed; 3 Tk 9 tests skipped on Tk 8.6.15.
- Both frozen applications initialized native Tk and verified Chinese paths,
  fresh-process recovery, exact 1,265-cell restoration, AI formulas, chart and hidden-sheet retention.
- Automated checks requested recalculation; they did not execute Excel calculation.

## Completed Windows Excel 2016 manual acceptance (user report, 2026-10-31)

- Mask → simulated AI edits → Compare → Restore; all 1,265 protected cells recovered, formulas recalculated and results checked.
- After a full Windows reboot, the same batch was recognized and restored correctly.
- Edited protected names were reported as conflicts and AI values retained.
- Edited record IDs stopped automatic restoration for affected rows; no identity guessing.
- The designated privacy sentinel count in SYNTHETIC_HR.masked.xlsx was zero.

This receipt concerns the fictional sample and tested configuration. It is not
a security certification or universal workbook-fidelity claim. Exact desktop OS
build/Excel bitness and repair-dialog observations were not separately recorded
in this receipt; no additional claim is made. No repeat manual test is requested.

## Distribution

Mac uses ad-hoc signing and is not notarized; Windows EXE is unsigned.
Windows 10/11 x64 and the official Microsoft VC++ x64 runtime are required.
Documentation repacks do not constitute a new Windows execution.
Publication requires a separate source-rights review and explicit approval.

## Final CSV compatibility removal (2026-10-31)

The unused CSV reader, RecoveryMaterial and format-dispatch/fallback loader
were removed. GUI, Compare, Restore, record identity and the encrypted V1
protocol remain unchanged. CSV and other non-.enc inputs are explicitly
rejected; invalid .enc ciphertext never falls back to a different format.
The old 304-core baseline loses three compatibility tests and gains four
rejection cases: the final core suite contains 305 tests.

Mac final candidate: 305 core and 144 isolated GUI tests passed; native frozen
Tk, Chinese paths, 1,265/1,265 exact restoration, formulas, chart and veryHidden
sheet passed. An encrypted fictional batch created by the previous V1 code was
also recovered by the new frozen app. Name conflicts retained the edited value
(1264 safe / 1 conflict); changed IDs blocked the affected row's 10 fields
(1255 safe / 10 confirmations).

Windows final candidate: 305 core and 141 isolated GUI tests passed; 3 Tk 9
tests were skipped on Tk 8.6.15. The native x64 frozen app passed real Tk,
Chinese paths, fresh-process .enc recovery, exact 1,265/1,265 restoration,
AI formulas, chart and hidden-sheet checks. The runtime is Python 3.12.10.
The source snapshot records native evidence in release/evidence/NATIVE_ACCEPTANCE.json. The prior Excel 2016 manual acceptance
remains historical evidence for the prior build; it is not relabelled as a
manual test of the rebuilt candidate. No repeat full manual acceptance is
required for this compatibility-only removal.
