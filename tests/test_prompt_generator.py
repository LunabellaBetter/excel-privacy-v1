"""Comprehensive tests for Checkpoint 7: Dynamic AI Prompt Generator.

Verifies:
1. text_only_prompt: only text rules, no numeric analysis rules
2. numeric_only_prompt: only numeric rules, no token rules
3. mixed_text_numeric_prompt: both sections present
4. multi_sheet_prompt: multiple sheets displayed independently
5. different_header_rows: custom header_row shown when != 1
6. different_record_identity_per_sheet: independent keys per sheet
7. composite_record_identity: formatted with ' + '
8. no_key_prompt_does_not_overpromise_restore: no overpromising when no key
9. duplicate_headers_show_column_letters: physical column disambiguation
10. blank_header_display_safe: [空表头] handled cleanly
11. numeric_prompt_allows_relative_analysis: ratio, ranking, percentage, trend
12. numeric_prompt_warns_against_absolute_thresholds: explicitly warns
13. prompt_recommends_new_suggestion_column_for_amount_changes: recommend new column over overwrite
14. prompt_warns_not_to_modify_tokens: preserve MSK token strings
15. prompt_allows_added_analysis_columns: explicit permission
16. prompt_allows_added_sheet: explicit permission
17. prompt_discourages_pivot_slicer_connections: pre-flight safety warning
18. prompt_does_not_include_scale_factor: no k / scale factor
19. prompt_does_not_include_run_id: no run_id
20. prompt_does_not_include_secret_key: no Fernet secret key
21. prompt_does_not_include_local_private_path: no private_keys directory
22. prompt_does_not_include_original_sensitive_values: no real values
23. real_task_token_not_embedded: no actual task token embedded
24. header_row_one_not_noisily_repeated: omitted when header_row == 1
25. header_row_non_one_is_shown: shown when header_row != 1
26. limited_precision_wording_is_present: precision disclaimer present
27. formula_recalculation_wording_is_accurate: dynamic formula recommended
28. prompt_is_plain_copyable_text: user-friendly and copyable
29. integration_with_mask_workbook_v1: returns ai_prompt in result dictionary
30. build_prompt_context_from_ledger: rebuilds cleanly from RecoveryLedger
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import sys
import pytest
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import privacy_handler.excel_privacy_tool as tool
from privacy_handler.prompt_generator import (
    PromptContext,
    PromptFieldInfo,
    PromptGeneratorSecurityError,
    PromptSheetScope,
    build_prompt_context_from_ledger,
    build_prompt_context_from_metadata,
    generate_ai_processing_prompt,
    validate_prompt_security,
)
from privacy_handler.table_scope import SheetProtectionScope


def test_01_text_only_prompt():
    """1. Text-only protection renders token rules but prunes numeric analysis section."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="UserList",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="B", header="姓名")],
                numeric_fields=[],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="用户ID")],
                has_record_identity=True,
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "文本隐私占位符处理规则" in prompt
    assert "可分析金额字段处理规则" not in prompt
    assert "MSK-XXXXXXXX-XXXXXXXX" in prompt
    assert "B｜姓名" in prompt
    assert "用户ID" in prompt


def test_02_numeric_only_prompt():
    """2. Numeric-only protection renders numeric analysis rules but prunes token section."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Finance",
                header_row=1,
                text_fields=[],
                numeric_fields=[PromptFieldInfo(column_letter="C", header="薪资")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="工号")],
                has_record_identity=True,
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "可分析数值保护字段处理规则" in prompt
    assert "文本隐私占位符处理规则" not in prompt
    assert "相对关系保护" in prompt
    assert "C｜薪资" in prompt


def test_03_mixed_text_numeric_prompt():
    """3. Mixed protection renders both text token and numeric analysis sections."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Payroll",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="B", header="员工姓名")],
                numeric_fields=[PromptFieldInfo(column_letter="D", header="基本工资")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="工号")],
                has_record_identity=True,
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "文本隐私占位符处理规则" in prompt
    assert "可分析数值保护字段处理规则" in prompt


def test_04_multi_sheet_prompt():
    """4. Multi-sheet prompt clearly delineates scopes per worksheet."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="人员历史",
                header_row=5,
                text_fields=[PromptFieldInfo(column_letter="B", header="姓名")],
                numeric_fields=[PromptFieldInfo(column_letter="D", header="基本工资")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="工号")],
                has_record_identity=True,
            ),
            PromptSheetScope(
                sheet_name="人员快照",
                header_row=1,
                text_fields=[],
                numeric_fields=[PromptFieldInfo(column_letter="C", header="当前薪资")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="工号")],
                has_record_identity=True,
            ),
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "【工作表：人员历史】" in prompt
    assert "【工作表：人员快照】" in prompt
    assert "表头位置：第 5 行" in prompt


def test_05_different_header_rows():
    """5. Non-row-1 header is explicitly shown; row 1 header is omitted."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="SheetA",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="A", header="ColA")],
            ),
            PromptSheetScope(
                sheet_name="SheetB",
                header_row=4,
                text_fields=[PromptFieldInfo(column_letter="A", header="ColB")],
            ),
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    # SheetA header_row=1 not shown
    assert "【工作表：SheetA】\n- 字段表头位置" not in prompt
    # SheetB header_row=4 shown
    assert "【工作表：SheetB】\n- 字段表头位置：第 4 行" in prompt


def test_06_different_record_identity_per_sheet():
    """6. Independent record identities displayed per worksheet."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Employees",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="B", header="Name")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="EmpID")],
                has_record_identity=True,
            ),
            PromptSheetScope(
                sheet_name="Orders",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="C", header="Customer")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="OrderID")],
                has_record_identity=True,
            ),
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "A｜EmpID" in prompt
    assert "A｜OrderID" in prompt


def test_07_composite_record_identity():
    """7. Composite record identity components formatted cleanly with ' + '."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="MonthlyLogs",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="C", header="Event")],
                record_identity_fields=[
                    PromptFieldInfo(column_letter="A", header="EmpID"),
                    PromptFieldInfo(column_letter="B", header="YearMonth"),
                ],
                has_record_identity=True,
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "记录关联列：A｜EmpID + B｜YearMonth" in prompt


def test_08_no_key_prompt_does_not_overpromise_restore():
    """8. Without record identity, prompt does not list record identity column."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="SummaryOnly",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="A", header="Item")],
                record_identity_fields=[],
                has_record_identity=False,
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "记录关联列" not in prompt


def test_09_duplicate_headers_show_column_letters():
    """9. Columns with identical headers show their respective physical column letters."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Comparison",
                header_row=1,
                numeric_fields=[
                    PromptFieldInfo(column_letter="C", header="金额"),
                    PromptFieldInfo(column_letter="D", header="金额"),
                ],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "C｜金额" in prompt
    assert "D｜金额" in prompt


def test_10_blank_header_display_safe():
    """10. Blank headers displayed safely as '[空表头]' with column letter."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Data",
                header_row=1,
                text_fields=[
                    PromptFieldInfo(column_letter="F", header="", is_blank=True),
                ],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "F｜[空表头]" in prompt


def test_11_numeric_prompt_allows_relative_analysis():
    """11. Prompt explicitly permits ranking, ratio, percentage, trend, and relative growth."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Amount")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "数值大小排序与相对排名" in prompt
    assert "结构占比分析" in prompt
    assert "同比、环比与相对增长率计算" in prompt
    assert "相对趋势分析" in prompt


def test_12_numeric_prompt_warns_against_absolute_thresholds():
    """12. Prompt explicitly warns against hardcoded absolute currency thresholds."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Amount")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "切勿依据固定真实数值阈值直接判定" in prompt
    assert "切勿将当前表格中的脱敏数字直接当作真实数值" in prompt


def test_13_prompt_recommends_new_suggestion_column_for_amount_changes():
    """13. Recommends adding a new suggestion column rather than overwriting protected source columns."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Salary")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "数值分析与建议方案原则" in prompt
    assert "优先新增“建议值 / 模拟值”列或新增分析 Sheet" in prompt
    assert "不要直接覆盖修改" in prompt
    assert "直接覆盖原始受保护单元格会触发本地还原冲突" in prompt


def test_14_prompt_warns_not_to_modify_tokens():
    """14. Warns AI not to edit, abbreviate, or guess real identities behind tokens."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                text_fields=[PromptFieldInfo(column_letter="A", header="Name")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "切勿改写、缩写、替换为自创编号或猜测真实身份" in prompt
    assert "完整保留该 Token 字符串" in prompt


def test_15_prompt_allows_added_analysis_columns():
    """15. Explicitly encourages adding new analysis columns."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="Sheet1", header_row=1)]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "允许在现有工作表中新增分析列" in prompt


def test_16_prompt_allows_added_sheet():
    """16. Explicitly encourages adding new analysis sheets."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="Sheet1", header_row=1)]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "允许新增工作表" in prompt


def test_17_prompt_discourages_pivot_slicer_connections():
    """17. Specifically warns against unsupported PivotTables, Slicers, and connections."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="Sheet1", header_row=1)]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "数据透视表 (PivotTable)" in prompt
    assert "切片器 (Slicer)" in prompt
    assert "外部数据连接 / Power Query" in prompt
    assert "自定义 XML 映射" in prompt


def test_18_prompt_does_not_include_scale_factor():
    """18. Strictly zero exposure of scale factor k or numerical multiplier in prompt."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Finance",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Cost")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "scale_factor" not in prompt
    assert "scale factor" not in prompt.lower()
    assert "倍率是" not in prompt
    assert "随机倍率" not in prompt


def test_19_prompt_does_not_include_run_id():
    """19. Security scan enforces that run_id is never present."""
    fake_run_id = "d87a2e58-45b6-4f47-8178-57d424b910ee"
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="Sheet1", header_row=1)]
    )
    prompt = generate_ai_processing_prompt(ctx, forbidden_values_to_check=[fake_run_id])
    assert fake_run_id not in prompt


def test_20_prompt_does_not_include_secret_key():
    """20. Security scan ensures secret key is blocked."""
    fake_key = "dGVzdF9mZXJuZXRfa2V5XzEyMzQ1Njc4OTA="
    with pytest.raises(PromptGeneratorSecurityError):
        validate_prompt_security(f"Here is secret: {fake_key}", forbidden_values=[fake_key])


def test_21_prompt_does_not_include_local_private_path():
    """21. Internal file paths (private_keys, recovery_ledger) trigger security error."""
    with pytest.raises(PromptGeneratorSecurityError):
        validate_prompt_security("Output is in /path/to/private_keys/data.enc")


def test_22_prompt_does_not_include_original_sensitive_values():
    """22. Original sensitive values are blocked if present in security check list."""
    orig_ssn = "110101199003072214"
    with pytest.raises(PromptGeneratorSecurityError):
        validate_prompt_security(f"SSN: {orig_ssn}", forbidden_values=[orig_ssn])


def test_23_real_task_token_not_embedded():
    """23. Real task token MSK-A7D2F914-X8F3K2M9 triggers security error while MSK-XXXXXXXX-XXXXXXXX is allowed."""
    # Placeholder example passes
    validate_prompt_security("Values look like MSK-XXXXXXXX-XXXXXXXX in cell")
    # Real token fails
    with pytest.raises(PromptGeneratorSecurityError):
        validate_prompt_security("Values look like MSK-A7D2F914-X8F3K2M9 in cell")


def test_24_header_row_one_not_noisily_repeated():
    """24. Standard header_row=1 is not redundantly printed."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="A", header="Name")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "表头位置：第 1 行" not in prompt


def test_25_header_row_non_one_is_shown():
    """25. Non-standard header_row=5 is clearly printed."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="CustomTable",
                header_row=5,
                text_fields=[PromptFieldInfo(column_letter="B", header="Name")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "字段表头位置：第 5 行" in prompt


def test_26_limited_precision_wording_is_present():
    """26. Prompt accurately notes that limited precision quantization may cause minor rounding differences."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Stats",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Val")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "有限精度量化可能产生极微小舍入误差" in prompt


def test_27_formula_recalculation_wording_is_accurate():
    """27. Formula recalculation note mentions that Excel will request recalculation upon opening."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Stats",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Val")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "真实数值阈值与参数化公式说明" in prompt
    assert "Excel 重新计算后，才得到真实判断结果" in prompt


def test_28_prompt_is_plain_copyable_text():
    """28. Prompt is clear, well-structured Chinese text suitable for direct clipboard copying."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Employees",
                header_row=1,
                text_fields=[PromptFieldInfo(column_letter="B", header="姓名")],
                numeric_fields=[PromptFieldInfo(column_letter="C", header="基本工资")],
                record_identity_fields=[PromptFieldInfo(column_letter="A", header="工号")],
                has_record_identity=True,
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert isinstance(prompt, str)
    assert len(prompt) > 200
    assert prompt.startswith("【Excel 分析与处理说明】")
    # Verify no unformatted brackets or developer jargon
    assert "protected_field_id" not in prompt
    assert "TableScope" not in prompt
    assert "CSPRNG" not in prompt
    assert "Fernet" not in prompt


def test_29_integration_with_mask_workbook_v1(tmp_path: Path):
    """29. mask_workbook_v1 automatically generates and returns ai_prompt in result dictionary."""
    src = tmp_path / "test_prompt_int.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Salary"])
    ws.append(["E01", "Alice", 10000])
    ws.append(["E02", "Bob", 20000])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        text_columns=["Name"],
        numeric_columns=["Salary"],
        record_key="ID",
        numeric_scale_factor=Decimal("2.0"),
    )

    assert "ai_prompt" in res
    prompt_str = res["ai_prompt"]
    assert "Name" in prompt_str
    assert "Salary" in prompt_str
    assert "ID" in prompt_str
    assert "可分析数值保护字段处理规则" in prompt_str
    assert "文本隐私占位符处理规则" in prompt_str


def test_30_build_prompt_context_from_ledger(tmp_path: Path):
    """30. PromptContext can be reconstructed directly from loaded RecoveryLedger."""
    src = tmp_path / "test_ledger_ctx.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Code", "Bonus"])
    ws.append(["C1", 500])
    wb.save(src)

    res = tool.mask_workbook_v1(
        input_path=src,
        output_dir=tmp_path / "safe",
        private_key_dir=tmp_path / "priv",
        header_row=1,
        numeric_columns=["Bonus"],
        record_key="Code",
    )

    ledger = res["ledger"]
    ctx = build_prompt_context_from_ledger(ledger)
    prompt = generate_ai_processing_prompt(ctx)
    assert "Bonus" in prompt
    assert "Code" in prompt
    assert "可分析数值保护字段处理规则" in prompt


# =========================================================================
# Checkpoint 7 Final Patch: Numeric Analysis Group Boundary Tests (31 - 42)
# =========================================================================


def test_31_single_numeric_group_allows_cross_field_relative_analysis():
    """31. When all fields belong to a single numeric group, state that they share the same ratio and can be compared relatively."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="SalarySheet",
                numeric_fields=[
                    PromptFieldInfo(column_letter="C", header="基本工资", numeric_analysis_group="grp1"),
                    PromptFieldInfo(column_letter="D", header="绩效奖金", numeric_analysis_group="grp1"),
                ],
                numeric_analysis_group="grp1",
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "以下保护数值使用同一比例保护" in prompt
    assert "仅在业务含义允许时进行彼此之间的相对分析" in prompt
    assert "跨组分析硬边界" not in prompt


def test_32_multiple_numeric_groups_warn_against_cross_group_comparison():
    """32. When multiple numeric groups exist, explicitly warn against cross-group ratio/amount comparison."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Finance",
                numeric_fields=[
                    PromptFieldInfo(column_letter="C", header="基本工资", numeric_analysis_group="salary_grp"),
                    PromptFieldInfo(column_letter="E", header="项目成本", numeric_analysis_group="cost_grp"),
                ],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "保护数值分析组划分（非常重要）" in prompt
    assert "【跨组分析硬边界】：不同保护组之间的比例尺度不保证一致" in prompt
    assert "请不要直接进行跨组数值大小、占比或比率比较" in prompt
    assert "切勿跨不同分析组直接计算数值比例、比率、跨组差值或大小横向比较" in prompt


def test_33_cross_sheet_same_numeric_group_is_presented_as_comparable():
    """33. Cross-sheet fields belonging to the same numeric group are grouped together as comparable."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="人员历史",
                numeric_fields=[PromptFieldInfo(column_letter="D", header="基本工资", numeric_analysis_group="payroll_group")],
                numeric_analysis_group="payroll_group",
            ),
            PromptSheetScope(
                sheet_name="人员快照",
                numeric_fields=[PromptFieldInfo(column_letter="F", header="当前薪资", numeric_analysis_group="payroll_group")],
                numeric_analysis_group="payroll_group",
            ),
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    # Both sheets share payroll_group, so only 1 group exists overall
    assert "以下保护数值使用同一比例保护" in prompt
    assert "人员历史｜D｜基本工资" in prompt
    assert "人员快照｜F｜当前薪资" in prompt
    assert "跨组分析硬边界" not in prompt


def test_34_same_header_different_groups_not_treated_as_same_group():
    """34. Same header names on different sheets with different groups are treated as separate groups."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="国内业务",
                numeric_fields=[PromptFieldInfo(column_letter="C", header="预算", numeric_analysis_group="domestic_grp")],
                numeric_analysis_group="domestic_grp",
            ),
            PromptSheetScope(
                sheet_name="海外业务",
                numeric_fields=[PromptFieldInfo(column_letter="C", header="预算", numeric_analysis_group="overseas_grp")],
                numeric_analysis_group="overseas_grp",
            ),
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "保护数值分析组 1" in prompt
    assert "保护数值分析组 2" in prompt
    assert "【跨组分析硬边界】" in prompt


def test_35_single_field_group_only_allows_within_field_relative_analysis():
    """35. A single-field group states it only allows within-field sorting/comparison, not cross-group comparison."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="DeptA",
                numeric_fields=[
                    PromptFieldInfo(column_letter="B", header="薪资1", numeric_analysis_group="g1"),
                    PromptFieldInfo(column_letter="C", header="薪资2", numeric_analysis_group="g1"),
                ],
            ),
            PromptSheetScope(
                sheet_name="DeptB",
                numeric_fields=[
                    PromptFieldInfo(column_letter="D", header="单项补贴", numeric_analysis_group="g2"),
                ],
            ),
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "该组仅含单个字段，仅允许字段内部排序与同字段不同记录间的相对比较，不可与其他组数值保护字段直接做比例比较" in prompt


def test_36_prompt_never_exposes_group_scale_factor():
    """36. Scale factors (k numbers) are strictly forbidden from entering the prompt text."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Finance",
                numeric_fields=[PromptFieldInfo(column_letter="C", header="Amount", numeric_analysis_group="grp1")],
                numeric_analysis_group="grp1",
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "scale_factor" not in prompt
    assert "scale factor" not in prompt.lower()
    assert "k =" not in prompt
    assert "倍率" not in prompt or "使用同一比例保护" in prompt  # Only allowed wording is '同一比例'


def test_37_prompt_does_not_use_internal_group_secret_metadata():
    """37. Internal secret IDs / UUIDs / hashes are not leaked into the prompt."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Finance",
                numeric_fields=[PromptFieldInfo(column_letter="C", header="Amount", numeric_analysis_group="secret_id_999")],
                numeric_analysis_group="secret_id_999",
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx, forbidden_values_to_check=["secret_id_999"])
    assert "secret_id_999" not in prompt


def test_38_amount_suggestion_prefers_new_column_or_new_sheet():
    """38. Prompt specifically guides AI to prefer adding new suggestion column or new analysis Sheet."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Salary")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "优先新增“建议值 / 模拟值”列或新增分析 Sheet" in prompt


def test_39_prompt_does_not_force_all_ai_output_next_to_source_column():
    """39. Prompt explicitly clarifies that AI outputs are not restricted to be right next to the source column."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="Sheet1", header_row=1)]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "并不限制所有 AI 成果都必须放在原字段旁边" in prompt
    assert "你可以自由选择新增分析列、新增汇总 Sheet、新增公式、新增图表或文字分析报告" in prompt


def test_40_absolute_threshold_formula_marked_provisional_until_restore():
    """40. Absolute threshold formula results are clearly marked provisional until restored by Excel."""
    ctx = PromptContext(
        sheets=[
            PromptSheetScope(
                sheet_name="Sheet1",
                numeric_fields=[PromptFieldInfo(column_letter="B", header="Salary")],
            )
        ]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "在当前脱敏文件中，基于真实数值阈值的公式计算结果可能暂时没有业务意义" in prompt
    assert "用户在本地安全恢复真实数值并由 Excel 重新计算后，才得到真实判断结果" in prompt
    assert "请勿在恢复前直接基于该中间结果形成确定性业务结论" in prompt


def test_41_real_token_is_blocked():
    """41. Real task tokens like MSK-A1B2C3D4-E5F6G7H8 are strictly rejected by security scan."""
    with pytest.raises(PromptGeneratorSecurityError):
        validate_prompt_security("Some token MSK-A1B2C3D4-E5F6G7H8 is leaked")


def test_42_placeholder_token_example_is_allowed():
    """42. Only placeholder token MSK-XXXXXXXX-XXXXXXXX is allowed by security scan."""
    # Should not raise
    validate_prompt_security("Placeholder format: MSK-XXXXXXXX-XXXXXXXX is opaque")


def test_43_prompt_does_not_claim_all_sensitive_data_is_protected():
    """43. Prompt opening strictly never overpromises that 'all sensitive data' or the 'whole workbook' is protected."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="S1", text_fields=[PromptFieldInfo(column_letter="A", header="Name")])]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "所有敏感信息" not in prompt
    assert "整个工作簿已完全匿名" not in prompt
    assert "全表无敏感数据" not in prompt
    assert "绝对安全" not in prompt


def test_44_prompt_mentions_selected_or_designated_protected_fields():
    """44. Prompt opening accurately refers to specified/selected protected fields."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="S1", text_fields=[PromptFieldInfo(column_letter="A", header="Name")])]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "本文件中指定的文本与数值保护字段已完成本地隐私保护处理" in prompt


def test_45_prompt_does_not_claim_strict_linear_ratio_preservation_after_quantization():
    """45. Prompt does not claim strict or 100% linear ratio invariance under decimal quantization."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="S1", numeric_fields=[PromptFieldInfo(column_letter="B", header="Amount")])]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "严格保持了相对大小与线性比例" not in prompt
    assert "比例严格完全不变" not in prompt
    assert "100% 保真" not in prompt


def test_46_prompt_mentions_possible_minor_precision_error():
    """46. Prompt explicitly warns that finite precision quantization may introduce minor rounding errors."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="S1", numeric_fields=[PromptFieldInfo(column_letter="B", header="Amount")])]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "有限精度处理可能产生微小舍入误差" in prompt
    assert "高保真保持比例关系" in prompt


def test_47_generic_edit_wording_is_neutral():
    """47. Generic prompt does not assume the workbook is about compensation."""
    ctx = PromptContext(
        sheets=[PromptSheetScope(sheet_name="Metrics", numeric_fields=[PromptFieldInfo(column_letter="B", header="Score")])]
    )
    prompt = generate_ai_processing_prompt(ctx)
    assert "数值调整建议、模拟方案、预测值或测算结果" in prompt
    assert "调薪建议" not in prompt
    assert "奖金调整方案" not in prompt


def test_48_formula_guidance_defaults_to_excel_2016_functions():
    """48. Formula guidance recommends Excel 2016-compatible functions by default."""
    ctx = PromptContext(sheets=[PromptSheetScope(sheet_name="Metrics")])
    prompt = generate_ai_processing_prompt(ctx)
    assert "为兼容 Excel 2016" in prompt
    assert "COUNTIF, SUMIF, VLOOKUP, INDEX + MATCH" in prompt
    assert "若使用 XLOOKUP 等高版本函数需确保用户运行环境支持" in prompt

