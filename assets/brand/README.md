# Excel Privacy 品牌资产台账 / Brand Asset Ledger

> **草案模板（DRAFT TEMPLATE）** — 本文件是 `assets/brand/README.md` 的**模板**，随 R3.1 Review Package 提供。
> 尚未写入任何仓库。§4 表格内的四个 SHA-256 占位标记由 `build_icons.sh` 在 **GENERATE 阶段**
> 用四项真实资产 SHA-256 自动替换后，才由 COMMIT 阶段提交。
> **不需要、也不允许手工修改正式仓库来补齐哈希。** Gate B′ 授权前不执行。
> （实现说明：占位标记采用双 @ 前缀的 SHA256 键名格式，本段说明刻意不写出字面标记，
> 以免 GENERATE 阶段的「占位标记无残留」断言把说明文字误判为未替换项。）

## 1. 资产清单（对应 `tools/stage_release.py` 的 8 个显式 BASE_FILES 路径）

| # | 路径 | 类别 | 来源 / 生成方式 |
|---|---|---|---|
| 1 | `BRAND_ASSETS.md` | 登记（文本） | R3.1 `drafts/BRAND_ASSETS.md` 逐字复制 |
| 2 | `assets/brand/README.md` | **生成（台账）** | 本模板 + §4 四项真实 SHA-256 自动填充（R3.1 C） |
| 3 | `assets/brand/excel-privacy-logo-master.png` | **生成（图像）** | Master 原图 `cp -p` 字节一致复制 |
| 4 | `assets/brand/excel-privacy-logo-1024.png` | **生成（图像）** | `sips -z 1024 1024` 等比缩放 |
| 5 | `assets/brand/icons/mac/ExcelPrivacy.icns` | **生成（图像）** | 10 槽 iconset → `iconutil -c icns` |
| 6 | `assets/brand/icons/win/ExcelPrivacy.ico` | **生成（图像）** | `scripts/build_ico.py`（7 条目：16–128 DIB + 256 PNG） |
| 7 | `assets/brand/scripts/build_icons.sh` | 登记（脚本） | R3.1 `scripts/build_icons.sh` 逐字复制 |
| 8 | `assets/brand/scripts/build_ico.py` | 登记（脚本） | R3.1 `scripts/build_ico.py` 逐字复制 |

**数量口径**：**4 项实际品牌图像资产**（第 3–6 行）+ **1 项生成台账**（第 2 行）+ **3 项登记文本/脚本**（第 1、7、8 行）= **8 个 BASE_FILES 路径**。
其中第 2 项虽属「登记路径」，但内容是**由模板自动生成**（不再逐字复制），因此模板 SHA 与最终 SHA 必然不同，见 §4。

> **不登记** `assets/brand/excel-privacy-logo.svg`：**未提供任何 SVG / AI / Figma 原始矢量源**。
> 登记一个不存在的文件会让 `stage_release.py` 的 `shutil.copyfile` 直接失败。不描摹、不自动矢量化。

## 2. Master 原图（唯一设计基准）

| 项 | 值 |
|---|---|
| 文件 | `logo-excel privacy.png`（保留在 `~/Downloads/自制logo/`，不移动、不覆盖、不重命名） |
| 尺寸 | 1254 × 1254 |
| 色彩 | PNG / 8-bit RGB（**无 alpha 通道**；R0 实测，非 RGBA） |
| SHA-256 | `0be01b184f5298ba73b5012fff09c18a76a9d149f0e6833baa12ff4be139b287` |
| 设计 | 暖象牙白背景、鼠尾草绿月牙、左侧白色 3×4 网格、右侧砂金色与陶土红 3×4 数据块 |

设计已由 Una 确认。**不得重新设计、改变比例、颜色、图形排列或重新绘制**。
本目录的派生资产只做等比缩放与无损容器转换；**不抠图、不去背、不调色、不描摹、不自动矢量化**。

## 3. 生成与复现

在隔离实施分支（Gate B′ 授权后）执行：

```bash
bash assets/brand/scripts/build_icons.sh \
     "<master.png>" "<repo-root>" "<r3.1-package-root>"
```

脚本分四阶段，**全部输出到临时暂存目录后才提交**：

1. **PREPARE**：把 4 个登记文件（含本 README **模板**）从第 3 个参数指定的 R3.1 包复制到暂存路径；
   对每个文件做「草案 SHA-256 ↔ 暂存 SHA-256」比对并打印台账；同时记录 **README 模板 SHA**。
   **不假设目标目录已存在构建脚本** —— 脚本与 `build_ico.py` 由本阶段显式复制。
2. **GENERATE**：校验 Master（SHA-256 + 1254×1254 + 8-bit RGB）→ 原样复制 →
   1024 PNG（含 1024×1024 尺寸复检）→ 10 槽 ICNS → 7 尺寸 ICO。
3. **GENERATE-README（R3.1 C）**：四项图像资产完成后，读取各自真实 SHA-256，
   把本模板 §4 的四个占位符**自动替换**，生成最终 README；计算 **README 最终 SHA**；
   断言模板 SHA ≠ 最终 SHA、占位符无残留、四个哈希与实物逐一相等；
   最后确认 8 个 BASE_FILES **全部**具有最终 SHA-256。**全程无手工编辑仓库**。
4. **COMMIT（R3.1 B）**：每个文件先写入**同目录的 run-unique 临时文件**并校验 SHA-256，
   再用 `ln` 硬链接**原子、不覆盖**地发布到最终路径，发布后复算 SHA-256。
   `cp` 若在部分写入后失败，只会留下可识别的临时文件，绝不污染正式资产路径；
   失败回滚删除本次临时文件与本次成功发布的文件，且**发布前二次校验 SHA，绝不删除他人文件**。
   回滚后同一目录可安全重跑。

工具链：macOS `sips` + `iconutil` + Python 标准库（`zlib` / `struct` / `hashlib` / `re`）。**不引入 Pillow / ImageMagick 等第三方依赖**。

## 4. SHA-256 台账（R3.1：由脚本自动填入真实值）

| 资产 | SHA-256 |
|---|---|
| `excel-privacy-logo-master.png` | `0be01b184f5298ba73b5012fff09c18a76a9d149f0e6833baa12ff4be139b287` |
| `excel-privacy-logo-1024.png` | `b0e8b324a331a8ea06a7b3d8fedfbc03f300359d578b0467f80d95fbaf332656` |
| `icons/mac/ExcelPrivacy.icns` | `422addb2d7a819b5051bd7e435bb17fb30bf6df855834503ad79a1e388603b79` |
| `icons/win/ExcelPrivacy.ico` | `e33b139677788fa454c8d8ebf23485d40a9ae89cae4d79ed04cd654f47fb9a83` |

**两个 SHA 必须区分，不得混为一项**：

| 名称 | 含义 | 用途 |
|---|---|---|
| `README_TEMPLATE_SHA256` | 本文件（§4 表格仍为四个待填占位标记）的 SHA-256 | 证明模板来源与草案一致 |
| `README_FINAL_SHA256` | 占位符被四项真实值替换后、实际提交进仓库的 README SHA-256 | 进入 `BASE_FILES` / `SOURCE_MANIFEST` 的最终值 |

README **自身**的 SHA-256 由脚本计算并打印，但**不写进 README 内部**（避免自指循环）。
四个图像资产的 SHA-256 由 GENERATE 阶段实物计算后写入；四个登记文本/脚本的 SHA-256
由 PREPARE/COMMIT 阶段计算。8 个 BASE_FILES 的最终 SHA-256 会由脚本逐项打印并可核对。
`release/SOURCE_MANIFEST.json` 由 `tools/stage_release.py` 重生成，是**发行侧的权威哈希清单**；
本台账是面向维护者的可读副本，两者不一致时以 manifest 为准并排查原因。

## 5. ICNS 反解检查 ≠ 实际视觉检查

| 检查 | 做什么 | 能证明什么 | 不能证明什么 |
|---|---|---|---|
| `iconutil -c iconset` 反解 | 把 ICNS 反解回 iconset，断言恰好 10 个槽位 | **容器结构完整、槽位齐全、可被系统解析** | 不证明颜色、清晰度、圆角裁切、月牙与网格是否美观正确 |
| **实际视觉检查** | 在 macOS Finder / Dock / 「显示简介」/ Quick Look / 关于窗口人工目视 | 图标在真实系统呈现下的可读性、圆角适配、缩放清晰度 | 需人工执行，脚本无法替代 |

**反解通过不得判为视觉验收 PASS。** 视觉检查属于 Gate B′ 的独立验收项，须留截图证据。
Windows 侧同理：`--verify-only` 只证明 ICO 容器与像素数据完整，不证明 EXE 文件图标、快捷方式、
运行中任务栏 / Alt-Tab 或 GUI 窗口标题栏显示正确；这些分别由 R1 §7 的双平台验收标准定义。

## 6. 权属与使用边界

- 本次确认的 Master 及其派生 PNG / ICNS / ICO 为 **LunaBella 原创品牌资产**，详见 `BRAND_ASSETS.md`。
- 品牌资产**不自动适用**软件 MIT `LICENSE`；软件 MIT 条款本身不变。
- 未来外部设计方提供的 SVG / AI / Figma 或其他资产，**必须单独确认权属与授权关系**，不自动纳入本声明。
- 第三方使用须取得单独书面授权，本台账不构成授权，也不扩大任何既有授权范围。

## 7. 相关文件

- 权属声明：`BRAND_ASSETS.md`
- 第三方组件许可：`THIRD_PARTY_NOTICES.md`（仅追加一段品牌说明，不改现有条款）
- 软件许可：`LICENSE`（MIT，保持原样）
- 独立发行白名单产物：`release/SOURCE_MANIFEST.json`
- 白名单登记脚本：`tools/stage_release.py` 的 `BASE_FILES`（8 个显式路径，**不使用 glob**）
