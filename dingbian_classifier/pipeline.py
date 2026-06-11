"""End-to-end classification pipeline."""

from __future__ import annotations

from pathlib import Path

from .classifier import classify_rows
from .coefficients import apply_manual_coefficients, fill_coefficients, prepare_coefficients
from .decomposition import (
    decompose_rolling_remark_rules,
    decompose_c6_heat_pump_dryer,
    decompose_composite_penguin_c6,
    decompose_rolling_final,
    decompose_skd,
    decompose_t7p7t5p5_dreame,
    decompose_t9p9,
    decompose_t9p9_dryer,
    decompose_t10p10,
    decompose_wave_basic,
    decompose_wave_final,
    write_extra_order_summary,
)
from .excel_repair import resave_with_excel_if_available
from .excel_io import build_output_path, copy_workbook, load_workbook_pair, move_auxiliary_sheets_after, read_main_table
from .formatting import format_main_sheet
from .history import backup_current_result, cleanup_output_results
from .logger import ProcessingLogger
from .material_description import (
    apply_manual_material_descriptions,
    fill_material_descriptions,
    prepare_material_descriptions,
)
from .reporter import write_log_sheet, write_results
from .rule_config import ensure_config_sheet, load_rules
from .sheet_metal import (
    apply_manual_sheet_metal_models,
    ensure_standard_units_column,
    fill_sheet_metal_models,
    prepare_sheet_metal,
)
from .sheet_detector import find_target_sheet


def set_active_sheet(workbook, sheet_name: str, logger: ProcessingLogger) -> None:
    workbook.active = workbook.sheetnames.index(sheet_name)
    logger.info(f"已设置默认打开工作表：{sheet_name}")


def run(
    input_path: Path,
    output_dir: Path,
    stage: str = "classify",
    coefficient_lookup: Path | None = None,
    sheet_metal_lookup: Path | None = None,
    material_description_lookup: Path | None = None,
) -> Path:
    logger = ProcessingLogger()
    input_path = input_path.resolve()
    output_dir = output_dir.resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")

    backup_current_result(input_path, output_dir, stage, logger)
    output_path = build_output_path(input_path, output_dir)
    copy_workbook(input_path, output_path, logger)

    formula_wb, values_wb = load_workbook_pair(output_path)
    original_sheet_names = set(formula_wb.sheetnames)
    target_sheet_name = find_target_sheet(formula_wb.sheetnames, logger)

    if stage == "reorder-sheets":
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        logger.info("已将系数/钣金型号相关辅助工作表移动到周排产明细表后面。")
        logger.info(f"即将保存处理结果：{output_path}")
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "format-main-sheet":
        format_main_sheet(formula_wb[target_sheet_name], logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-skd":
        decompose_skd(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-rolling-remarks":
        decompose_rolling_remark_rules(formula_wb, values_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-t7p7t5p5-dreame":
        decompose_t7p7t5p5_dreame(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-t9p9":
        decompose_t9p9(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-t9p9-dryer":
        decompose_t9p9_dryer(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-t10p10":
        decompose_t10p10(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-c6-heat-pump-dryer":
        decompose_c6_heat_pump_dryer(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-composite-penguin-c6":
        decompose_composite_penguin_c6(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-rolling-final":
        decompose_rolling_final(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-wave-basic":
        decompose_wave_basic(formula_wb, values_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-wave-final":
        decompose_wave_final(formula_wb, values_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "decompose-extra-summary":
        write_extra_order_summary(formula_wb, values_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        set_active_sheet(formula_wb, target_sheet_name, logger)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "prepare-coefficients":
        prepare_result = prepare_coefficients(formula_wb, values_wb, target_sheet_name, logger)
        if prepare_result.coefficient_missing_rows:
            logger.warning("已生成“系数补充”工作表，流程暂停；请查询系数后再运行 fill-coefficients。")
        else:
            logger.info("未发现系数为 #N/A 的行，可以继续后续分类流程。")
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "fill-coefficients":
        if coefficient_lookup is None:
            raise ValueError("fill-coefficients 阶段必须提供 --coefficient-lookup。")
        fill_coefficients(formula_wb, values_wb, target_sheet_name, coefficient_lookup.resolve(), logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "apply-manual-coefficients":
        apply_manual_coefficients(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "prepare-sheet-metal":
        sheet_metal_result = prepare_sheet_metal(formula_wb, values_wb, target_sheet_name, logger)
        if sheet_metal_result.sheet_metal_missing_rows:
            logger.warning("已生成“钣金型号补充”工作表，流程暂停；请查询钣金型号后提供查询规则。")
        else:
            logger.info("未发现钣金型号为 #N/A 的行，可以继续后续流程。")
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "fill-sheet-metal":
        if sheet_metal_lookup is None:
            raise ValueError("fill-sheet-metal 阶段必须提供 --sheet-metal-lookup。")
        fill_sheet_metal_models(
            formula_wb,
            values_wb,
            target_sheet_name,
            sheet_metal_lookup.resolve(),
            logger,
        )
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "apply-manual-sheet-metal":
        apply_manual_sheet_metal_models(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "prepare-material-description":
        result = prepare_material_descriptions(formula_wb, values_wb, target_sheet_name, logger)
        if result.missing_rows:
            logger.warning("已生成“物料描述补充”工作表，流程暂停；请查询物料描述后再运行 fill-material-description。")
        else:
            logger.info("未发现物料描述为空白或 #N/A 的行，可以继续后续流程。")
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "fill-material-description":
        if material_description_lookup is None:
            raise ValueError("fill-material-description 阶段必须提供 --material-description-lookup。")
        fill_material_descriptions(
            formula_wb,
            values_wb,
            target_sheet_name,
            material_description_lookup.resolve(),
            logger,
        )
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "apply-manual-material-description":
        apply_manual_material_descriptions(formula_wb, target_sheet_name, logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    if stage == "prepare-standard-units":
        ensure_standard_units_column(formula_wb[target_sheet_name], logger)
        logger.info(f"即将保存处理结果：{output_path}")
        move_auxiliary_sheets_after(formula_wb, target_sheet_name)
        write_log_sheet(formula_wb, logger)
        formula_wb.save(output_path)
        resave_with_excel_if_available(output_path, logger)
        cleanup_output_results(output_dir, output_path, logger)
        logger.info(f"处理完成：{output_path}")
        return output_path

    headers, rows, _raw_rows = read_main_table(values_wb[target_sheet_name])
    logger.info(f"主数据表读取完成：表头 {len(headers)} 列，数据 {len(rows)} 行。")

    ensure_config_sheet(formula_wb, logger)
    rules = load_rules(formula_wb, headers, logger)
    result = classify_rows(rows, rules)
    logger.info(
        f"分类执行完成：生成分类 {len(result.categories)} 个，未分类 {len(result.unmatched)} 行。"
    )
    if result.unmatched:
        logger.warning("存在未分类数据，请查看“未分类数据”工作表并补充分类规则。")

    logger.info(f"即将保存处理结果：{output_path}")
    move_auxiliary_sheets_after(formula_wb, target_sheet_name)
    write_results(formula_wb, headers, result, original_sheet_names, logger)
    formula_wb.save(output_path)
    resave_with_excel_if_available(output_path, logger)
    cleanup_output_results(output_dir, output_path, logger)
    logger.info(f"处理完成：{output_path}")
    return output_path
