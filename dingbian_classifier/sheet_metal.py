"""Standard-unit and sheet-metal-model preparation helpers."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .data_cleaner import deduplicate_headers, normalize_header
from .excel_io import remove_sheet_if_exists
from .exceptions import MissingRequiredFieldsError
from .logger import ProcessingLogger
from .reporter import style_header

STANDARD_UNITS_FIELD = "标台数"
SHEET_METAL_SUPPLEMENT_SHEET = "钣金型号补充"
SHEET_METAL_LOOKUP_SHEET = "钣金型号查询表"


@dataclass
class SheetMetalPrepareResult:
    standard_units_formulas: int
    sheet_metal_missing_rows: int


@dataclass
class SheetMetalFillResult:
    filled_rows: int
    remaining_rows: int
    lookup_rows: int


@dataclass
class ManualSheetMetalApplyResult:
    applied_rows: int
    remaining_rows: int


def prepare_sheet_metal(
    workbook: Workbook,
    values_workbook: Workbook,
    target_sheet_name: str,
    logger: ProcessingLogger,
) -> SheetMetalPrepareResult:
    formula_sheet = workbook[target_sheet_name]
    values_sheet = values_workbook[target_sheet_name]

    missing_count = create_sheet_metal_supplement_sheet(
        workbook,
        formula_sheet,
        values_sheet,
        logger,
    )
    return SheetMetalPrepareResult(
        standard_units_formulas=0,
        sheet_metal_missing_rows=missing_count,
    )


def ensure_standard_units_column(sheet: Worksheet, logger: ProcessingLogger) -> int:
    headers = _header_map(sheet)
    _require_columns(headers, ["系数", "订单数"], "主数据表")

    if STANDARD_UNITS_FIELD not in headers:
        insert_col = headers["订单数"]
        sheet.insert_cols(insert_col, 1)
        sheet.cell(1, insert_col).value = STANDARD_UNITS_FIELD
        logger.info(f"已在“订单数”列前新增“{STANDARD_UNITS_FIELD}”列。")
    else:
        logger.info(f"检测到已存在“{STANDARD_UNITS_FIELD}”列，将刷新公式和格式。")

    headers = _header_map(sheet)
    _require_columns(headers, ["系数", STANDARD_UNITS_FIELD, "订单数"], "主数据表")
    coefficient_col = headers["系数"]
    standard_col = headers[STANDARD_UNITS_FIELD]
    order_col = headers["订单数"]

    coefficient_letter = get_column_letter(coefficient_col)
    order_letter = get_column_letter(order_col)
    for row_index in range(2, sheet.max_row + 1):
        cell = sheet.cell(row_index, standard_col)
        cell.value = f"={coefficient_letter}{row_index}*{order_letter}{row_index}"
        cell.number_format = "0.00"
    sheet.cell(1, standard_col).number_format = "General"

    logger.info(f"已刷新“{STANDARD_UNITS_FIELD}”公式：{max(sheet.max_row - 1, 0)} 行。")
    return max(sheet.max_row - 1, 0)


def refresh_daily_output_formulas(sheet: Worksheet, logger: ProcessingLogger) -> None:
    headers = _header_map(sheet)
    required = ["日产", "订单数", "线体", "基本开始日期"]
    if any(name not in headers for name in required):
        logger.warning("未刷新“日产”公式：缺少 日产/订单数/线体/基本开始日期 中的一个字段。")
        return

    daily_col = headers["日产"]
    order_col = headers["订单数"]
    line_col = headers["线体"]
    date_col = headers["基本开始日期"]
    order_letter = get_column_letter(order_col)
    line_letter = get_column_letter(line_col)
    date_letter = get_column_letter(date_col)

    for row_index in range(2, sheet.max_row + 1):
        sheet.cell(row_index, daily_col).value = (
            f"=SUMIFS({order_letter}:{order_letter},"
            f"{line_letter}:{line_letter},{line_letter}{row_index},"
            f"{date_letter}:{date_letter},{date_letter}{row_index})"
        )
    logger.info("已同步刷新“日产”列公式，确保继续引用移动后的“订单数”列。")


def create_sheet_metal_supplement_sheet(
    workbook: Workbook,
    formula_sheet: Worksheet,
    values_sheet: Worksheet,
    logger: ProcessingLogger,
) -> int:
    values_headers = _header_map(values_sheet)
    _require_columns(values_headers, ["钣金型号"], "主数据表")
    sheet_metal_col = values_headers["钣金型号"]

    remove_sheet_if_exists(workbook, SHEET_METAL_SUPPLEMENT_SHEET)
    supplement = workbook.create_sheet(SHEET_METAL_SUPPLEMENT_SHEET)
    output_headers = ["原始行号"] + [
        normalize_header(cell.value) or f"空列{index}"
        for index, cell in enumerate(formula_sheet[1], start=1)
    ]
    supplement.append(output_headers)

    missing_count = 0
    for row_index in range(2, values_sheet.max_row + 1):
        value = values_sheet.cell(row_index, sheet_metal_col).value
        if not _is_na(value):
            continue
        missing_count += 1
        row_values = [
            formula_sheet.cell(row_index, col_index).value
            for col_index in range(1, formula_sheet.max_column + 1)
        ]
        supplement.append([row_index] + row_values)

    supplement.freeze_panes = "A2"
    if supplement.max_row > 1:
        supplement.auto_filter.ref = supplement.dimensions
    style_header(supplement)
    _fit_columns(supplement)
    logger.info(f"钣金型号为 #N/A 的行已摘取到“{SHEET_METAL_SUPPLEMENT_SHEET}”：{missing_count} 行。")
    return missing_count


def fill_sheet_metal_models(
    workbook: Workbook,
    values_workbook: Workbook,
    target_sheet_name: str,
    lookup_path,
    logger: ProcessingLogger,
) -> SheetMetalFillResult:
    lookup_pairs = read_sheet_metal_lookup(lookup_path, logger)
    lookup_map: dict[Any, Any] = {}
    conflicts = 0
    for code, model in lookup_pairs:
        if _is_blank(code) or _is_blank(model) or _is_na(model):
            continue
        if code in lookup_map and lookup_map[code] != model:
            conflicts += 1
            continue
        lookup_map[code] = model

    import_sheet_metal_lookup_sheet(workbook, lookup_pairs, logger)

    formula_sheet = workbook[target_sheet_name]
    values_sheet = values_workbook[target_sheet_name]
    headers = _header_map(values_sheet)
    _require_columns(headers, ["钣金型号", "物料编码"], "主数据表")
    sheet_metal_col = headers["钣金型号"]
    code_col = headers["物料编码"]

    remaining_rows: list[int] = []
    filled_rows = 0
    for row_index in range(2, values_sheet.max_row + 1):
        current_value = values_sheet.cell(row_index, sheet_metal_col).value
        if not _is_na(current_value):
            continue
        code = values_sheet.cell(row_index, code_col).value
        matched_model = lookup_map.get(code)
        if matched_model is None:
            remaining_rows.append(row_index)
            continue
        formula_sheet.cell(row_index, sheet_metal_col).value = matched_model
        filled_rows += 1

    refresh_sheet_metal_supplement_sheet(workbook, formula_sheet, remaining_rows, logger)
    if conflicts:
        logger.warning(f"钣金型号查询表存在同一物料编码对应多个钣金型号的情况，已保留首次匹配：{conflicts} 条。")
    logger.info(
        f"钣金型号已按物料编码回填：{filled_rows} 行；仍待补充 {len(remaining_rows)} 行。"
    )
    return SheetMetalFillResult(
        filled_rows=filled_rows,
        remaining_rows=len(remaining_rows),
        lookup_rows=len(lookup_pairs),
    )


def apply_manual_sheet_metal_models(
    workbook: Workbook,
    target_sheet_name: str,
    logger: ProcessingLogger,
) -> ManualSheetMetalApplyResult:
    if SHEET_METAL_SUPPLEMENT_SHEET not in workbook.sheetnames:
        logger.info(f"未找到“{SHEET_METAL_SUPPLEMENT_SHEET}”，无需手工回填钣金型号。")
        return ManualSheetMetalApplyResult(applied_rows=0, remaining_rows=0)

    source_sheet = workbook[SHEET_METAL_SUPPLEMENT_SHEET]
    formula_sheet = workbook[target_sheet_name]
    source_headers = _header_map(source_sheet)
    main_headers = _header_map(formula_sheet)
    _require_columns(source_headers, ["原始行号", "钣金型号"], SHEET_METAL_SUPPLEMENT_SHEET)
    _require_columns(main_headers, ["钣金型号"], "主数据表")

    row_no_col = source_headers["原始行号"]
    source_model_col = source_headers["钣金型号"]
    main_model_col = main_headers["钣金型号"]
    remaining_rows: list[int] = []
    applied_rows = 0

    for row_index in range(2, source_sheet.max_row + 1):
        original_row = source_sheet.cell(row_index, row_no_col).value
        model = source_sheet.cell(row_index, source_model_col).value
        if _is_blank(original_row):
            continue
        try:
            original_row_number = int(original_row)
        except (TypeError, ValueError):
            logger.warning(f"{SHEET_METAL_SUPPLEMENT_SHEET} 第 {row_index} 行原始行号无效，已跳过：{original_row}")
            continue

        if _is_valid_manual_model(model):
            formula_sheet.cell(original_row_number, main_model_col).value = model
            applied_rows += 1
        else:
            remaining_rows.append(original_row_number)

    refresh_sheet_metal_supplement_sheet(workbook, formula_sheet, remaining_rows, logger)
    logger.info(f"已回填手工钣金型号：{applied_rows} 行；剩余暂不处理 {len(remaining_rows)} 行。")
    return ManualSheetMetalApplyResult(applied_rows=applied_rows, remaining_rows=len(remaining_rows))


def read_sheet_metal_lookup(lookup_path, logger: ProcessingLogger) -> list[tuple[Any, Any]]:
    import openpyxl

    if not lookup_path.exists():
        raise FileNotFoundError(f"钣金型号查询表不存在：{lookup_path}")
    lookup_wb = openpyxl.load_workbook(lookup_path, data_only=True)
    source_sheet = lookup_wb[lookup_wb.sheetnames[0]]
    headers = _header_map(source_sheet)
    _require_columns(headers, ["物料编码", "钣金型号"], "钣金型号查询表")
    code_col = headers["物料编码"]
    model_col = headers["钣金型号"]

    rows: list[tuple[Any, Any]] = []
    for row_index in range(2, source_sheet.max_row + 1):
        code = source_sheet.cell(row_index, code_col).value
        model = source_sheet.cell(row_index, model_col).value
        if _is_blank(code):
            continue
        rows.append((code, model))
    logger.info(f"已读取钣金型号查询表：{len(rows)} 条。")
    return rows


def import_sheet_metal_lookup_sheet(
    workbook: Workbook,
    lookup_pairs: list[tuple[Any, Any]],
    logger: ProcessingLogger,
) -> None:
    remove_sheet_if_exists(workbook, SHEET_METAL_LOOKUP_SHEET)
    sheet = workbook.create_sheet(SHEET_METAL_LOOKUP_SHEET)
    sheet.append(["物料编码", "钣金型号"])
    for code, model in lookup_pairs:
        sheet.append([code, model])
    sheet.freeze_panes = "A2"
    if sheet.max_row > 1:
        sheet.auto_filter.ref = sheet.dimensions
    style_header(sheet)
    _fit_columns(sheet)
    logger.info(f"已导入“{SHEET_METAL_LOOKUP_SHEET}”：{len(lookup_pairs)} 条。")


def refresh_sheet_metal_supplement_sheet(
    workbook: Workbook,
    formula_sheet: Worksheet,
    row_numbers: list[int],
    logger: ProcessingLogger,
) -> None:
    remove_sheet_if_exists(workbook, SHEET_METAL_SUPPLEMENT_SHEET)
    supplement = workbook.create_sheet(SHEET_METAL_SUPPLEMENT_SHEET)
    output_headers = ["原始行号"] + [
        normalize_header(cell.value) or f"空列{index}"
        for index, cell in enumerate(formula_sheet[1], start=1)
    ]
    supplement.append(output_headers)
    for row_index in row_numbers:
        row_values = [
            formula_sheet.cell(row_index, col_index).value
            for col_index in range(1, formula_sheet.max_column + 1)
        ]
        supplement.append([row_index] + row_values)

    supplement.freeze_panes = "A2"
    if supplement.max_row > 1:
        supplement.auto_filter.ref = supplement.dimensions
    style_header(supplement)
    _fit_columns(supplement)
    logger.info(f"已刷新“{SHEET_METAL_SUPPLEMENT_SHEET}”：剩余待补充 {len(row_numbers)} 行。")


def _header_map(sheet: Worksheet) -> dict[str, int]:
    raw_headers = [normalize_header(cell.value) for cell in sheet[1]]
    headers = deduplicate_headers(raw_headers)
    return {header: index + 1 for index, header in enumerate(headers)}


def _require_columns(header_map: dict[str, int], columns: list[str], context: str) -> None:
    missing = [column for column in columns if column not in header_map]
    if missing:
        raise MissingRequiredFieldsError(f"{context}缺少关键字段：" + "、".join(missing))


def _is_na(value: Any) -> bool:
    return value is not None and str(value).strip().upper() in {"#N/A", "#NA", "N/A"}


def _is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def _is_valid_manual_model(value: Any) -> bool:
    if _is_blank(value) or _is_na(value):
        return False
    if isinstance(value, str) and value.strip().startswith("="):
        return False
    return True


def _fit_columns(sheet: Worksheet) -> None:
    for column_cells in sheet.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 45)
