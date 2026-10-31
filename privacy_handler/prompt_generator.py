"""Dynamic AI Prompt Generator for V1 privacy protocol.

Generates safe, clear, plain-language Excel processing instructions that users can
directly copy and paste to AI (such as ChatGPT, Claude, Codex, etc.).

Security Invariants:
1. Pure Public Metadata:
   Only uses sheet names, column display names, column letters, header row info,
   and safe general instructions.
2. Zero Secrets Leakage:
   Strictly never accesses or exposes scale factor k, secret keys, recovery ledger paths,
   run_id, FP8 fingerprints, real task token values, or original plaintext data.
3. Plain Copyable Text:
   No developer jargon (no Ledger, FP8, Fernet, CSPRNG, protected_field_id, TableScope, etc.).
4. Multi-Sheet & Heterogeneous Scopes:
   Renders clear, sheet-by-sheet scopes with independent header rows, protection types,
   and Record Identity columns.
5. Dynamic Section Pruning:
   Only outputs sections relevant to the active protection types (Text-only, Numeric-only, Mixed).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import Any

from privacy_handler.identity import TOKEN_SEARCH_REGEX
from privacy_handler.recovery_ledger import RecoveryLedger


class PromptGeneratorSecurityError(Exception):
    """Raised when generated prompt violates privacy or security constraints."""


@dataclass
class PromptFieldInfo:
    """Safe column descriptor for prompt generation."""

    column_letter: str
    header: str
    raw_header: str | None = None
    is_blank: bool = False
    numeric_analysis_group: str | None = None

    @property
    def display_name(self) -> str:
        """Formatted display: 'A｜员工工号' or 'D｜[空表头]'."""
        clean_header = self.header.strip() if self.header else ""
        if not clean_header or clean_header == "[空表头]" or self.is_blank:
            clean_header = "[空表头]"
        return f"{self.column_letter}｜{clean_header}"


@dataclass
class PromptSheetScope:
    """Per-sheet configuration for prompt generation."""

    sheet_name: str
    header_row: int = 1
    text_fields: list[PromptFieldInfo] = field(default_factory=list)
    numeric_fields: list[PromptFieldInfo] = field(default_factory=list)
    record_identity_fields: list[PromptFieldInfo] = field(default_factory=list)
    has_record_identity: bool = False
    numeric_analysis_group: str | None = None

    @property
    def has_text_protection(self) -> bool:
        return len(self.text_fields) > 0

    @property
    def has_numeric_protection(self) -> bool:
        return len(self.numeric_fields) > 0


@dataclass
class PromptContext:
    """Consolidated public metadata for AI prompt generation."""

    sheets: list[PromptSheetScope] = field(default_factory=list)
    protocol_version: str = "1.0"

    @property
    def has_any_text_protection(self) -> bool:
        return any(s.has_text_protection for s in self.sheets)

    @property
    def has_any_numeric_protection(self) -> bool:
        return any(s.has_numeric_protection for s in self.sheets)

    @property
    def has_any_record_identity(self) -> bool:
        return any(s.has_record_identity for s in self.sheets)

    def get_all_numeric_fields_by_group(self) -> dict[str, list[tuple[str, PromptFieldInfo]]]:
        """Group all numeric fields across sheets by their effective numeric_analysis_group.

        Returns:
            dict mapping group_key -> list of (sheet_name, field_info)
        """
        groups: dict[str, list[tuple[str, PromptFieldInfo]]] = {}
        for s in self.sheets:
            for nf in s.numeric_fields:
                g_key = nf.numeric_analysis_group or s.numeric_analysis_group or f"__sheet_{s.sheet_name}"
                groups.setdefault(g_key, []).append((s.sheet_name, nf))
        return groups


def build_prompt_context_from_metadata(
    sheet_scopes_meta: list[dict[str, Any]],
    protected_fields: list[Any],
    confirmed_key_fields: list[str] | str | None = None,
) -> PromptContext:
    """Build PromptContext from public non-sensitive configuration/metadata."""
    # Group protected fields by sheet_name
    pf_by_sheet: dict[str, list[Any]] = {}
    for pf in protected_fields:
        s_name = getattr(pf, "sheet_name", None) or (pf.get("sheet_name") if isinstance(pf, dict) else "") or ""
        pf_by_sheet.setdefault(s_name, []).append(pf)

    # Normalize confirmed_key_fields
    global_keys: list[str] = []
    if confirmed_key_fields:
        if isinstance(confirmed_key_fields, str):
            global_keys = [k.strip() for k in confirmed_key_fields.split("+") if k.strip()]
        elif isinstance(confirmed_key_fields, (list, tuple)):
            for item in confirmed_key_fields:
                if isinstance(item, str):
                    for sub in item.split("+"):
                        if sub.strip():
                            global_keys.append(sub.strip())

    prompt_sheets: list[PromptSheetScope] = []

    for scope_dict in sheet_scopes_meta:
        sheet_name = scope_dict.get("sheet_name", "Sheet1")
        header_row = int(scope_dict.get("header_row", 1))
        selected_cols = scope_dict.get("selected_columns", [])
        sheet_num_group = scope_dict.get("numeric_analysis_group")

        # Map column letter and header from selected_columns
        col_map_by_letter: dict[str, dict[str, Any]] = {}
        col_map_by_header: dict[str, dict[str, Any]] = {}
        for col in selected_cols:
            c_letter = col.get("column_letter", "")
            c_header = col.get("header", "")
            if c_letter:
                col_map_by_letter[c_letter] = col
            if c_header:
                col_map_by_header[c_header] = col

        # Determine key fields for this sheet
        raw_key = scope_dict.get("record_key")
        sheet_keys: list[str] = []
        if raw_key:
            if isinstance(raw_key, str):
                sheet_keys = [k.strip() for k in raw_key.split(",") if k.strip()]
            elif isinstance(raw_key, (list, tuple)):
                sheet_keys = [str(k).strip() for k in raw_key if str(k).strip()]
        else:
            sheet_keys = list(global_keys)

        sheet_pfs = pf_by_sheet.get(sheet_name, [])
        if not sheet_pfs and "" in pf_by_sheet and len(sheet_scopes_meta) == 1:
            sheet_pfs = pf_by_sheet[""]

        text_pfs: list[PromptFieldInfo] = []
        numeric_pfs: list[PromptFieldInfo] = []

        for pf in sheet_pfs:
            p_type = getattr(pf, "protection_type", None) or (pf.get("protection_type") if isinstance(pf, dict) else "")
            c_letter = getattr(pf, "source_column_letter", None) or (pf.get("source_column_letter") if isinstance(pf, dict) else "")
            l_field = getattr(pf, "logical_field", None) or (pf.get("logical_field") if isinstance(pf, dict) else "")
            raw_h = getattr(pf, "raw_header", None) or (pf.get("raw_header") if isinstance(pf, dict) else None)
            d_header = getattr(pf, "display_header", None) or (pf.get("display_header") if isinstance(pf, dict) else l_field)
            pf_num_group = getattr(pf, "numeric_analysis_group", None) or (pf.get("numeric_analysis_group") if isinstance(pf, dict) else None)
            eff_num_group = pf_num_group or sheet_num_group

            # Match column descriptor for accurate letter and blank status
            col_desc = col_map_by_letter.get(c_letter) or col_map_by_header.get(l_field)
            if col_desc:
                c_letter = col_desc.get("column_letter", c_letter)
                raw_h = col_desc.get("raw_header", raw_h)
                d_header = col_desc.get("header", d_header)

            is_blank = raw_h is None or len(str(raw_h).strip()) == 0

            field_info = PromptFieldInfo(
                column_letter=c_letter or "A",
                header=d_header or l_field,
                raw_header=str(raw_h) if raw_h is not None else None,
                is_blank=is_blank,
                numeric_analysis_group=eff_num_group,
            )

            if p_type == "text_token":
                text_pfs.append(field_info)
            elif p_type == "numeric_scaled":
                numeric_pfs.append(field_info)

        # Build Record Identity field infos
        rkey_infos: list[PromptFieldInfo] = []
        for k_item in sheet_keys:
            # k_item can be letter (e.g. "A") or header name (e.g. "员工工号")
            col_desc = col_map_by_letter.get(k_item) or col_map_by_header.get(k_item)
            if col_desc:
                c_letter = col_desc.get("column_letter", "")
                d_head = col_desc.get("header", k_item)
                raw_h = col_desc.get("raw_header")
                is_blank = raw_h is None or len(str(raw_h).strip()) == 0
                rkey_infos.append(
                    PromptFieldInfo(
                        column_letter=c_letter or k_item,
                        header=d_head,
                        raw_header=str(raw_h) if raw_h is not None else None,
                        is_blank=is_blank,
                    )
                )
            else:
                rkey_infos.append(
                    PromptFieldInfo(
                        column_letter="",
                        header=k_item,
                        raw_header=k_item,
                        is_blank=False,
                    )
                )

        prompt_sheets.append(
            PromptSheetScope(
                sheet_name=sheet_name,
                header_row=header_row,
                text_fields=text_pfs,
                numeric_fields=numeric_pfs,
                record_identity_fields=rkey_infos,
                has_record_identity=len(rkey_infos) > 0,
                numeric_analysis_group=sheet_num_group,
            )
        )

    return PromptContext(sheets=prompt_sheets)


def build_prompt_context_from_ledger(ledger: RecoveryLedger) -> PromptContext:
    """Build PromptContext strictly from RecoveryLedger's configuration and metadata (zero secrets)."""
    cfg = ledger.configuration
    sheet_scopes_meta = cfg.sheet_scopes or []
    if not sheet_scopes_meta and ledger.protected_fields:
        # Synthesize fallback scope for single sheet legacy ledger
        unique_sheets = sorted(list({pf.sheet_name or "Sheet1" for pf in ledger.protected_fields}))
        for sname in unique_sheets:
            s_pfs = [pf for pf in ledger.protected_fields if (pf.sheet_name or "Sheet1") == sname]
            h_row = s_pfs[0].source_header_row or 1 if s_pfs else 1
            sheet_scopes_meta.append({
                "sheet_name": sname,
                "header_row": h_row,
                "selected_columns": [],
            })

    return build_prompt_context_from_metadata(
        sheet_scopes_meta=sheet_scopes_meta,
        protected_fields=ledger.protected_fields,
        confirmed_key_fields=cfg.confirmed_key_field,
    )


def validate_prompt_security(
    prompt_text: str,
    forbidden_values: list[str] | None = None,
) -> None:
    """Scan generated prompt for forbidden leakages (run_id, scale factors, private paths, actual tokens)."""
    # 1. Check for real token patterns (allow placeholder example MSK-XXXXXXXX-XXXXXXXX)
    for m in TOKEN_SEARCH_REGEX.finditer(prompt_text):
        token_str = m.group(0)
        # Placeholder example is explicitly allowed
        if token_str == "MSK-XXXXXXXX-XXXXXXXX":
            continue
        raise PromptGeneratorSecurityError(
            f"Prompt contains actual task token: '{token_str}'"
        )

    # 2. Check for private system keywords and overpromising claims
    forbidden_patterns = [
        r"private_keys",
        r"\.recovery_ledger",
        r"\.mask_secret\.key",
        r"Fernet",
        r"scale_factor\s*=",
        r"倍率是",
        r"k\s*=\s*\d",
        r"所有敏感信息均已",
        r"整个工作簿已完全匿名",
        r"全表无敏感数据",
        r"绝对安全",
    ]
    for pat in forbidden_patterns:
        if re.search(pat, prompt_text, re.IGNORECASE):
            raise PromptGeneratorSecurityError(
                f"Prompt contains forbidden internal leakage matching pattern: {pat}"
            )

    # 3. Check caller-supplied forbidden values (e.g. real run_id, secret key, absolute path)
    if forbidden_values:
        for val in forbidden_values:
            if val and isinstance(val, str) and len(val.strip()) >= 4:
                if val.strip() in prompt_text:
                    raise PromptGeneratorSecurityError(
                        f"Prompt contains forbidden sensitive value: '{val[:8]}...'"
                    )


def generate_ai_processing_prompt(
    context: PromptContext,
    forbidden_values_to_check: list[str] | None = None,
) -> str:
    """Generate plain-language, user-friendly Excel processing instructions for AI."""
    lines: list[list[str]] = []

    # Title & Opening
    lines.append(["【Excel 分析与处理说明】\n"])
    lines.append([
        "本文件中指定的文本与数值保护字段已完成本地隐私保护处理。",
        "请直接基于本文件的当前结构进行分析与处理，并在完成后直接返回修改后的 Excel 文件。\n",
    ])

    # Section 1: Sheets and Protected Fields
    sec_idx = 1
    lines.append([f"一、需要处理的工作表与字段配置"])
    for s in context.sheets:
        sheet_block = [f"\n【工作表：{s.sheet_name}】"]
        if s.header_row != 1:
            sheet_block.append(f"- 字段表头位置：第 {s.header_row} 行")

        # Record Identity
        if s.has_record_identity:
            key_names = [f.display_name if f.column_letter else f.header for f in s.record_identity_fields]
            sheet_block.append(f"- 记录关联列：{' + '.join(key_names)}")
            sheet_block.append("  （用于保持记录对应关系，请勿删除、重命名该列，亦不可单独对该列执行局部排序）")

        # Text fields
        if s.has_text_protection:
            sheet_block.append("- 文本隐私占位字段：")
            for f in s.text_fields:
                sheet_block.append(f"  * {f.display_name}")

        # Numeric fields
        if s.has_numeric_protection:
            sheet_block.append("- 可分析数值保护字段：")
            for f in s.numeric_fields:
                sheet_block.append(f"  * {f.display_name}")

        lines.append(sheet_block)

    sec_idx += 1

    # Section 2: Text Protection Rules (Conditional)
    if context.has_any_text_protection:
        sec_num_str = ["一", "二", "三", "四", "五", "六", "七", "八"][sec_idx - 1]
        text_block = [
            f"\n{sec_num_str}、文本隐私占位符处理规则",
            "1. 格式特征：文本保护字段中的值均为形如 `MSK-XXXXXXXX-XXXXXXXX` 的隐私占位符。",
            "2. 保持原样：请在处理过程中完整保留该 Token 字符串，切勿改写、缩写、替换为自创编号或猜测真实身份。",
            "3. 允许操作：可正常用于关联匹配、分组统计、计数、去重与整行排序。",
            "4. 还原保障：用户后续会在本地安全环境中将占位符自动恢复为真实业务文本。",
        ]
        lines.append(text_block)
        sec_idx += 1

    # Section 3: Numeric Protection Rules (Conditional)
    if context.has_any_numeric_protection:
        sec_num_str = ["一", "二", "三", "四", "五", "六", "七", "八"][sec_idx - 1]
        num_block = [f"\n{sec_num_str}、可分析数值保护字段处理规则"]

        # Analyze numeric analysis groups across all sheets
        groups_dict = context.get_all_numeric_fields_by_group()
        num_group_count = len(groups_dict)

        if num_group_count <= 1:
            # Single numeric group
            num_block.append("1. 相对关系保护：")
            num_block.append("   - 以下保护数值使用同一比例保护，可保持相对大小，并高保真保持比例关系；有限精度处理可能产生微小舍入误差，仅在业务含义允许时进行彼此之间的相对分析。")
            all_fields: list[str] = []
            for g_items in groups_dict.values():
                for s_name, f_info in g_items:
                    all_fields.append(f"{s_name}｜{f_info.display_name}")
            if all_fields:
                num_block.append(f"   - 包含字段：{'、'.join(all_fields)}")
        else:
            # Multiple numeric groups
            num_block.append("1. 保护数值分析组划分（非常重要）：")
            group_idx = 1
            for g_key, g_items in groups_dict.items():
                fields_str = "、".join(f"{s_name}｜{f_info.display_name}" for s_name, f_info in g_items)
                num_block.append(f"   - 【保护数值分析组 {group_idx}】：{fields_str}")
                if len(g_items) > 1:
                    num_block.append("     * 该组内字段经过同一比例保护，在业务含义允许时可以进行组内比例、占比、相对大小比较。")
                else:
                    num_block.append("     * 该组仅含单个字段，仅允许字段内部排序与同字段不同记录间的相对比较，不可与其他组数值保护字段直接做比例比较。")
                group_idx += 1
            num_block.append("   - 【跨组分析硬边界】：不同保护组之间的比例尺度不保证一致，请不要直接进行跨组数值大小、占比或比率比较。")

        num_block.extend([
            "2. 允许的分析操作：",
            "   - 同一分析组内部的数值大小排序与相对排名",
            "   - 同一分析组内部的结构占比分析（例如在业务含义允许时计算各项占该组总数值的百分比）",
            "   - 同一分析组内部的同比、环比与相对增长率计算",
            "   - 相对趋势分析与变化趋势图表",
            "   - 相关性分析（注意：有限精度量化可能产生极微小舍入误差）",
            "3. 严禁的分析操作：",
            "   - 切勿将当前表格中的脱敏数字直接当作真实数值进行绝对解读",
            "   - 切勿跨不同分析组直接计算数值比例、比率、跨组差值或大小横向比较",
            "   - 切勿依据固定真实数值阈值直接判定（例如禁止直接写死：'数值 > 10,000 即为高标准'）",
            "   - 切勿基于当前数值进行固定真实数值区间分桶",
            "4. 真实数值阈值与参数化公式说明：",
            "   - 若业务必须按真实数值门槛进行测算或分档，允许建立参数单元格并在公式中引用；",
            "   - 【注意】：在当前脱敏文件中，基于真实数值阈值的公式计算结果可能暂时没有业务意义；",
            "   - 用户在本地安全恢复真实数值并由 Excel 重新计算后，才得到真实判断结果。请勿在恢复前直接基于该中间结果形成确定性业务结论。",
        ])
        lines.append(num_block)
        sec_idx += 1

    # Section 4: Editing Guidelines
    sec_num_str = ["一", "二", "三", "四", "五", "六", "七", "八"][sec_idx - 1]
    edit_block = [
        f"\n{sec_num_str}、表格编辑与建议方案原则",
        "1. 数值分析与建议方案原则（非常重要）：",
        "   - 若需要提出数值调整建议、模拟方案、预测值或测算结果，请【优先新增“建议值 / 模拟值”列或新增分析 Sheet】，而【不要直接覆盖修改】原始受保护的数值或文本列；",
        "   - 直接覆盖原始受保护单元格会触发本地还原冲突，需要用户人工介入审核；",
        "   - 并不限制所有 AI 成果都必须放在原字段旁边，你可以自由选择新增分析列、新增汇总 Sheet、新增公式、新增图表或文字分析报告，只要不直接覆盖破坏受保护源字段即可。",
        "2. 保持整行数据对齐：",
        "   - 允许对整个工作表按行进行全局排序或筛选；",
        "   - 切勿单独对某一列进行局部重新排序，以免打乱行与行之间的业务关联。",
        "3. 允许的扩展操作：",
        "   - 允许在现有工作表中新增分析列、统计行、说明列；",
        "   - 允许新增工作表（例如新增《分析汇总》、《可视化看板》Sheet）；",
        "   - 允许插入原生 Excel 基础图表（折线图、柱状图、饼图等）；",
        "   - 允许编写原生工作表公式（为兼容 Excel 2016，推荐使用 SUM, AVERAGE, IF, COUNTIF, SUMIF, VLOOKUP, INDEX + MATCH 等通用函数；若使用 XLOOKUP 等高版本函数需确保用户运行环境支持）。",
        "4. 请避免使用的高级对象：",
        "   - 为确保回传文件能顺利进行本地还原，请勿新建以下高级对象：数据透视表 (PivotTable)、切片器 (Slicer)、外部数据连接 / Power Query、自定义 XML 映射。",
    ]
    lines.append(edit_block)
    sec_idx += 1

    # Section 5: Output Requirements
    sec_num_str = ["一", "二", "三", "四", "五", "六", "七", "八"][sec_idx - 1]
    out_block = [
        f"\n{sec_num_str}、输出交付要求",
        "1. 保持原有工作表名称及原有关键结构完好；",
        "2. 所有公式保持动态可编辑状态，无需写死为静态数值；",
        "3. 完成分析后，请直接输出修改后的完整 Excel 文件供用户下载回传。",
    ]
    lines.append(out_block)

    # Flatten and build final string
    prompt_content = "\n".join("\n".join(b) for b in lines).strip() + "\n"

    # Security scan
    validate_prompt_security(prompt_content, forbidden_values=forbidden_values_to_check)

    return prompt_content
