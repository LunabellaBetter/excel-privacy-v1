# Excel Privacy Tool

开发者：LunaBella

本地 Excel 字段保护与 AI 编辑后安全恢复工具。Python/Tkinter 界面，支持
macOS Apple Silicon 和 Windows x64。

**版本：V1.0.0。** 已完成双平台原生自动化，以及用户报告的 Windows Excel 2016 完整流程、公式重算、整机重启恢复与冲突验收。验证范围和限制见 `release/ACCEPTANCE.md`；正式下载以 GitHub Release 为准。

仅支持未加密的 `.xlsx`。V1 禁止 `.xlsm`、VBA（含改名为 xlsx）、旧版 xls、
透视表／缓存、切片器、外部连接／链接、custom XML、嵌入对象与 ActiveX。

## 使用流程

1. 备份原始工作簿，选择文件，逐表设置表头行。
2. 选择非空、唯一且稳定的记录识别列；必要时联合多列。
3. 选择文本保护、数值保护或不保护。识别列也需要按敏感性设置保护。
4. 按需设置数值可比组；不同表默认隔离，跨表共享必须明确配置。
5. 确认每张参与表，开始脱敏。检查整个输出文件后，将脱敏 xlsx 和本批次
   AI 说明交给 AI。原件、账本、密钥及恢复文件不得交给外部 AI。
6. 选择 AI 返回文件，检查修改，再恢复真实数据。阅读恢复数量及未完成项。
7. 在 Excel 中打开恢复文件、重算公式，并人工核对业务结果。

记录识别列遭修改、删除或出现歧义时，相关数据保留 AI 返回值并要求核对。
批次短参考号不是唯一安全凭证；缺少可信批次材料时停止自动恢复。
数值保护保留某些相对关系，不能保证抵抗已知值推断，也不是统计匿名化。

## 安装和本地运行

macOS arm64：解压发行 ZIP，把 `.app` 移入本地目录运行。macOS 应用仅使用临时签名、未公证；
若系统阻止打开，可在核对来源及 SHA-256 后使用系统“隐私与安全性”中的打开选项。
不要全局关闭 Gatekeeper。Windows x64：应用使用免安装 ZIP，须完整解压，不能只复制 exe。
Windows EXE 未签名。需要 Windows 10/11 的系统组件及 Microsoft Visual C++ 2015–2022 x64 运行库。
本包不再分发 Microsoft DLL；如果本机缺少运行库，请使用
[微软官方运行库](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170)，
安装完成后再启动应用。受管理设备请使用获准的软件安装方式。

独立应用将本地资料保存在 macOS 的 `~/Library/Application Support/Excel Privacy Tool/`
或 Windows 的 `%LOCALAPPDATA%\ExcelPrivacyTool\`。保留并安全备份整个恢复环境。
源码运行时默认使用工程内数据目录；不应把这些数据目录上传至 GitHub。

源码开发需要 Python 3.12 与可用的 Tkinter：

```sh
python -m pip install -r requirements.txt -r requirements-build.txt -r requirements-test.txt
python privacy_app.py --verify-self
python privacy_app.py
```

## 已知限制

- 仅保护明确选择的字段。未选列、备注、批注、公式字符串、图片、隐藏表及文件名
  可能仍含敏感信息，交付 AI 前必须检查；不承诺整份文件完全匿名。
- 工具不能计算公式。重算标记不等于 Excel 已完成重算。
- 不承诺全部 Excel 对象无损。先使用纯虚构样例验收，保留原件。
- 老批次若账本身份快照不一致，不会猜测修复；应保留旧环境，并从原始本地备份
  在新版本重新脱敏。不得混用不同批次恢复资料。
- 平台构建与真实 Excel 2016 兼容性验收分别记录，见 release/ACCEPTANCE.md。
- TXT、Markdown、DOCX、PDF、OCR 不属于 V1。

## 来源与许可

历史来源包括
kwen1510/excel-anonymiser-skill；其代码没有被本项目的 MIT 许可重新授权。
独立发行采用文件白名单，替换继承的辅助实现，并排除历史 CLI、Skill 模板及旧样例。
V1 工作流和安全模块另行维护。详见 release/SOURCE_PROVENANCE.md。
独立快照采用 MIT，第三方组件遵循各自许可，见 THIRD_PARTY_NOTICES.md。

公开源代码和发行包以实际 GitHub Release 页面为准。请核对版本、校验值及已知限制；未经确认的候选版本不应作为正式发布版本。

- V1 仅接受 `.enc` 加密恢复账本；已移除旧 CSV 恢复兼容和解密失败后的格式回退。既有 V1 加密账本继续使用原协议。
