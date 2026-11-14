# Excel Privacy Tool V1.0.1

Maintainer: LunaBella. Updated 2026-11-14.

## V1.0.1 (2026-11-14)
- 新增 LunaBella 原创品牌图标：由确认的 Master 原图（1254×1254）等比重缩放与格式转换派生 Master/1024 PNG、macOS ICNS（10 槽位）、Windows ICO（7 条目：16/24/32/48/64/128 DIB + 256 PNG）。
- 版本号由 1.0.0 升至 1.0.1：macOS `CFBundleShortVersionString/CFBundleVersion` 与 Windows 文件版本均更新为 (1,0,1,0)；Windows EXE 嵌入图标与版本信息资源。
- Windows 原生构建（GitHub Actions windows-latest）增加 Gate B′ 自动化校验：ICO 容器/PNG 像素、PE 资源读回、PyInstaller 保留 7 条目图标、文件版本 1.0.1.0。
- macOS `.app` 现在携带自定义 ICNS 图标（Finder / Dock / 简介 / Quick Look 已实机验证）。
- 功能与安全模块相对 V1.0.0 **无变更**；未继承私有库与公开库之间的 18 处内容差异，已保留公开 V1.0.0 的既有修复（仅接受 `.enc`、移除 CSV 恢复兼容等）。
- macOS `bundle_identifier` **身份决策待 Una 批准**：候选规范恢复 V1.0.0 既定 `io.lunabella.excelprivacy` 并保留 `NSHumanReadableCopyright`；当前已构建的测试 Mac 制品为 `io.unaprivacy.excel` 且缺少版权字段，若批准恢复则需重建 Mac 制品（DMG/ZIP 哈希随之变更）。V1.0.0 发布记录已明确该身份为既定值。
- 品牌权属声明 `BRAND_ASSETS.md` / `assets/brand/README.md` 仍为**草案**，待 Una 确认，不作为正式法律声明发布。
- 已知：macOS 应用临时签名、未公证；Windows EXE 未签名；均按要求不含 Microsoft DLL。

---

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
