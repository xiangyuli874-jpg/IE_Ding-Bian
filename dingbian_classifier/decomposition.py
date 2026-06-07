"""Production order decomposition rules and summary sheet."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .data_cleaner import deduplicate_headers, normalize_header
from .excel_io import remove_sheet_if_exists
from .exceptions import MissingRequiredFieldsError
from .logger import ProcessingLogger

DETAIL_SHEET_NAME = "排单分解表明细"
WAVE_LINES = {"B线", "B线夜", "C线", "C线夜"}
DECOMPOSE_COLUMNS = ["基本开始日期", "备注", "类型"]

SKD_MAIN_FILL = "00B050"
SKD_NORMAL_TYPE_FILL = "92D050"
SKD_DRY_TYPE_FILL = "F4CCCC"
CKD_MAIN_FILL = "FFC000"
CKD_NORMAL_TYPE_FILL = "FFD966"
CKD_DRUM_TYPE_FILL = "F4B183"
SAMSUNG_MAIN_FILL = "9DC3E6"
DOUBLE_DRUM_MAIN_FILL = "A9D18E"
T7P7_MAIN_FILL = "7030A0"
T7P7_WASH_TYPE_FILL = "D9EAD3"
T7P7_DRY_TYPE_FILL = "D9D2E9"
T7P7_WASH_TYPE = "T7/P7/T5/P5/追觅单洗"
T7P7_DRY_TYPE = "T7/P7/T5/P5/追觅烘干"
T9P9_MAIN_FILL = "00B0F0"
T9P9_WASH_TYPE_FILL = "C9DAF8"
T9P9_DRY_TYPE_FILL = "FCE4D6"
T9P9_WASH_TYPE = "T9/P9单洗"
T9P9_DRY_TYPE = "T9/P9烘干"

ROLLING_LABELS = [
    "SKD烘干",
    "SKD总数(包含烘干)",
    "CKD",
    "CKD含筒部装",
    "三星",
    "双滚筒",
    T7P7_DRY_TYPE,
    T7P7_WASH_TYPE,
    T9P9_WASH_TYPE,
    T9P9_DRY_TYPE,
    "T10/P10洗衣机",
    "T10/P10干衣机",
    "C6热泵干衣机",
    "T9/P9干衣机",
    "复式烘干",
    "复式单洗",
    "企鹅干衣机",
    "企鹅洗衣机",
    "C6单洗",
    "C6Q10烘干",
    "普通烘干",
    "内销",
    "外销",
]
WAVE_LABELS = [
    "CKD",
    "LG铁皮外销",
    "塑料内销",
    "P7P9",
    "SKD",
    "内销铁皮变频",
    "外销普通变频",
    "外销铁皮",
    "内销铁皮",
]


@dataclass
class DecomposeSkdResult:
    rolling_total: Decimal
    wave_total: Decimal
    skd_total: Decimal
    skd_dry: Decimal
    skd_rows: int
    skd_dry_rows: int


@dataclass
class DecomposeRollingRemarkResult:
    rolling_total: Decimal
    wave_total: Decimal
    category_totals: dict[str, Decimal] = field(default_factory=dict)
    category_rows: dict[str, int] = field(default_factory=dict)


@dataclass
class DecomposeT7P7T5P5DreameResult:
    rolling_total: Decimal
    wave_total: Decimal
    category_totals: dict[str, Decimal] = field(default_factory=dict)
    category_rows: dict[str, int] = field(default_factory=dict)
    skipped_dryer_rows: int = 0
    skipped_unknown_code_rows: int = 0


@dataclass
class DecomposeT9P9Result:
    rolling_total: Decimal
    wave_total: Decimal
    category_totals: dict[str, Decimal] = field(default_factory=dict)
    category_rows: dict[str, int] = field(default_factory=dict)
    skipped_dryer_rows: int = 0
    skipped_unknown_code_rows: int = 0
    skipped_remark_only_rows: int = 0


def decompose_skd(workbook: Workbook, main_sheet_name: str, logger: ProcessingLogger) -> DecomposeSkdResult:
    sheet = workbook[main_sheet_name]
    headers = _header_map(sheet)
    _require_columns(headers, ["线体", "订单数", "备注", "类型", *DECOMPOSE_COLUMNS], "主数据表")

    rolling_total = Decimal("0")
    wave_total = Decimal("0")
    skd_total = Decimal("0")
    skd_dry = Decimal("0")
    skd_rows = 0
    skd_dry_rows = 0

    for row_index in range(2, sheet.max_row + 1):
        line = str(sheet.cell(row_index, headers["线体"]).value or "").strip()
        order_qty = _to_decimal(sheet.cell(row_index, headers["订单数"]).value)
        if _is_wave_line(line):
            wave_total += order_qty
            continue

        rolling_total += order_qty
        if not _is_uncolored_for_decomposition(sheet, row_index, headers):
            continue

        remark = str(sheet.cell(row_index, headers["备注"]).value or "")
        if "SKD" not in remark.upper():
            continue

        is_dry = "烘干" in remark
        type_name = "SKD烘干" if is_dry else "SKD"
        _apply_skd_style(sheet, row_index, headers, type_name, is_dry)

        skd_total += order_qty
        skd_rows += 1
        if is_dry:
            skd_dry += order_qty
            skd_dry_rows += 1

    summary = _collect_decomposition_summary(sheet, headers)
    write_decomposition_detail_sheet(workbook, main_sheet_name, summary, logger)
    logger.info(
        f"SKD排单分解完成：命中 {skd_rows} 行，SKD总数 {skd_total}，SKD烘干 {skd_dry}。"
    )
    return DecomposeSkdResult(
        rolling_total=rolling_total,
        wave_total=wave_total,
        skd_total=skd_total,
        skd_dry=skd_dry,
        skd_rows=skd_rows,
        skd_dry_rows=skd_dry_rows,
    )


def decompose_rolling_remark_rules(
    workbook: Workbook,
    values_workbook: Workbook,
    main_sheet_name: str,
    logger: ProcessingLogger,
) -> DecomposeRollingRemarkResult:
    """Run rolling-line remark rules after SKD: CKD, Samsung, and double-drum."""
    sheet = workbook[main_sheet_name]
    values_sheet = values_workbook[main_sheet_name]
    headers = _header_map(sheet)
    _require_columns(headers, ["线体", "订单数", "备注", "类型", "系数", *DECOMPOSE_COLUMNS], "主数据表")

    category_totals: dict[str, Decimal] = {"CKD": Decimal("0"), "CKD含筒部装": Decimal("0"), "三星": Decimal("0"), "双滚筒": Decimal("0")}
    category_rows: dict[str, int] = {"CKD": 0, "CKD含筒部装": 0, "三星": 0, "双滚筒": 0}

    for row_index in range(2, sheet.max_row + 1):
        line = str(sheet.cell(row_index, headers["线体"]).value or "").strip()
        if _is_wave_line(line) or not _is_uncolored_for_decomposition(sheet, row_index, headers):
            continue

        remark = str(sheet.cell(row_index, headers["备注"]).value or "")
        remark_upper = remark.upper()
        order_qty = _to_decimal(sheet.cell(row_index, headers["订单数"]).value)

        if "CKD" in remark_upper:
            coefficient = _to_decimal(values_sheet.cell(row_index, headers["系数"]).value)
            type_name = "CKD含筒部装" if coefficient > Decimal("1.6") else "CKD"
            type_fill = CKD_DRUM_TYPE_FILL if type_name == "CKD含筒部装" else CKD_NORMAL_TYPE_FILL
            _apply_decomposition_style(sheet, row_index, headers, type_name, CKD_MAIN_FILL, type_fill)
        elif "三星" in remark:
            type_name = "三星"
            _apply_decomposition_style(sheet, row_index, headers, type_name, SAMSUNG_MAIN_FILL, SAMSUNG_MAIN_FILL)
        elif "双滚筒" in remark:
            type_name = "双滚筒"
            _apply_decomposition_style(sheet, row_index, headers, type_name, DOUBLE_DRUM_MAIN_FILL, DOUBLE_DRUM_MAIN_FILL)
        else:
            continue

        category_totals[type_name] += order_qty
        category_rows[type_name] += 1

    summary = _collect_decomposition_summary(sheet, headers)
    write_decomposition_detail_sheet(workbook, main_sheet_name, summary, logger)
    logger.info(
        "滚筒排单分解规则完成："
        f"CKD {category_rows['CKD']} 行/{category_totals['CKD']}，"
        f"CKD含筒部装 {category_rows['CKD含筒部装']} 行/{category_totals['CKD含筒部装']}，"
        f"三星 {category_rows['三星']} 行/{category_totals['三星']}，"
        f"双滚筒 {category_rows['双滚筒']} 行/{category_totals['双滚筒']}。"
    )
    return DecomposeRollingRemarkResult(
        rolling_total=summary["rolling_total"],
        wave_total=summary["wave_total"],
        category_totals=category_totals,
        category_rows=category_rows,
    )


def decompose_t7p7t5p5_dreame(
    workbook: Workbook,
    main_sheet_name: str,
    logger: ProcessingLogger,
) -> DecomposeT7P7T5P5DreameResult:
    sheet = workbook[main_sheet_name]
    headers = _header_map(sheet)
    _require_columns(
        headers,
        ["线体", "订单数", "物料编码", "物料描述", "类型", *DECOMPOSE_COLUMNS],
        "主数据表",
    )

    category_totals: dict[str, Decimal] = {T7P7_DRY_TYPE: Decimal("0"), T7P7_WASH_TYPE: Decimal("0")}
    category_rows: dict[str, int] = {T7P7_DRY_TYPE: 0, T7P7_WASH_TYPE: 0}
    skipped_dryer_rows = 0
    skipped_unknown_code_rows = 0

    for row_index in range(2, sheet.max_row + 1):
        line = str(sheet.cell(row_index, headers["线体"]).value or "").strip()
        if _is_wave_line(line) or not _is_uncolored_for_decomposition(sheet, row_index, headers):
            continue

        description = str(sheet.cell(row_index, headers["物料描述"]).value or "")
        if not _is_t7p7t5p5_dreame_description(description):
            continue

        product_kind = _product_kind_from_material_code(sheet.cell(row_index, headers["物料编码"]).value)
        if product_kind == "dryer":
            skipped_dryer_rows += 1
            continue
        if product_kind == "unknown":
            skipped_unknown_code_rows += 1
            continue

        type_name = T7P7_DRY_TYPE if product_kind == "dry" else T7P7_WASH_TYPE
        type_fill = T7P7_DRY_TYPE_FILL if product_kind == "dry" else T7P7_WASH_TYPE_FILL
        _apply_decomposition_style(sheet, row_index, headers, type_name, T7P7_MAIN_FILL, type_fill)

        order_qty = _to_decimal(sheet.cell(row_index, headers["订单数"]).value)
        category_totals[type_name] += order_qty
        category_rows[type_name] += 1

    summary = _collect_decomposition_summary(sheet, headers)
    write_decomposition_detail_sheet(workbook, main_sheet_name, summary, logger)
    logger.info(
        "T7/P7/T5/P5/追觅排单分解完成："
        f"{T7P7_DRY_TYPE} {category_rows[T7P7_DRY_TYPE]} 行/{category_totals[T7P7_DRY_TYPE]}，"
        f"{T7P7_WASH_TYPE} {category_rows[T7P7_WASH_TYPE]} 行/{category_totals[T7P7_WASH_TYPE]}，"
        f"跳过605干衣机 {skipped_dryer_rows} 行，"
        f"跳过无法识别编码 {skipped_unknown_code_rows} 行。"
    )
    return DecomposeT7P7T5P5DreameResult(
        rolling_total=summary["rolling_total"],
        wave_total=summary["wave_total"],
        category_totals=category_totals,
        category_rows=category_rows,
        skipped_dryer_rows=skipped_dryer_rows,
        skipped_unknown_code_rows=skipped_unknown_code_rows,
    )


def decompose_t9p9(
    workbook: Workbook,
    main_sheet_name: str,
    logger: ProcessingLogger,
) -> DecomposeT9P9Result:
    sheet = workbook[main_sheet_name]
    headers = _header_map(sheet)
    _require_columns(
        headers,
        ["线体", "订单数", "备注", "物料编码", "物料描述", "类型", *DECOMPOSE_COLUMNS],
        "主数据表",
    )

    category_totals: dict[str, Decimal] = {T9P9_DRY_TYPE: Decimal("0"), T9P9_WASH_TYPE: Decimal("0")}
    category_rows: dict[str, int] = {T9P9_DRY_TYPE: 0, T9P9_WASH_TYPE: 0}
    skipped_dryer_rows = 0
    skipped_unknown_code_rows = 0
    skipped_remark_only_rows = 0

    for row_index in range(2, sheet.max_row + 1):
        line = str(sheet.cell(row_index, headers["线体"]).value or "").strip()
        if _is_wave_line(line) or not _is_uncolored_for_decomposition(sheet, row_index, headers):
            continue

        remark = str(sheet.cell(row_index, headers["备注"]).value or "")
        description = str(sheet.cell(row_index, headers["物料描述"]).value or "")
        has_remark_candidate = "T9" in remark.upper() or "P9" in remark.upper()
        has_description_candidate = _is_t9p9_description(description)
        if not has_remark_candidate and not has_description_candidate:
            continue
        if not has_description_candidate:
            skipped_remark_only_rows += 1
            continue

        product_kind = _product_kind_from_material_code(sheet.cell(row_index, headers["物料编码"]).value)
        if product_kind == "dryer":
            skipped_dryer_rows += 1
            continue
        if product_kind == "unknown":
            skipped_unknown_code_rows += 1
            continue

        type_name = T9P9_DRY_TYPE if product_kind == "dry" else T9P9_WASH_TYPE
        type_fill = T9P9_DRY_TYPE_FILL if product_kind == "dry" else T9P9_WASH_TYPE_FILL
        _apply_decomposition_style(sheet, row_index, headers, type_name, T9P9_MAIN_FILL, type_fill)

        order_qty = _to_decimal(sheet.cell(row_index, headers["订单数"]).value)
        category_totals[type_name] += order_qty
        category_rows[type_name] += 1

    summary = _collect_decomposition_summary(sheet, headers)
    write_decomposition_detail_sheet(workbook, main_sheet_name, summary, logger)
    logger.info(
        "T9/P9排单分解完成："
        f"{T9P9_DRY_TYPE} {category_rows[T9P9_DRY_TYPE]} 行/{category_totals[T9P9_DRY_TYPE]}，"
        f"{T9P9_WASH_TYPE} {category_rows[T9P9_WASH_TYPE]} 行/{category_totals[T9P9_WASH_TYPE]}，"
        f"跳过605干衣机 {skipped_dryer_rows} 行，"
        f"跳过无法识别编码 {skipped_unknown_code_rows} 行，"
        f"跳过仅备注命中但物料描述前段未命中 {skipped_remark_only_rows} 行。"
    )
    return DecomposeT9P9Result(
        rolling_total=summary["rolling_total"],
        wave_total=summary["wave_total"],
        category_totals=category_totals,
        category_rows=category_rows,
        skipped_dryer_rows=skipped_dryer_rows,
        skipped_unknown_code_rows=skipped_unknown_code_rows,
        skipped_remark_only_rows=skipped_remark_only_rows,
    )


def write_decomposition_detail_sheet(
    workbook: Workbook,
    main_sheet_name: str,
    summary: dict[str, Any],
    logger: ProcessingLogger,
) -> None:
    remove_sheet_if_exists(workbook, DETAIL_SHEET_NAME)
    sheet = workbook.create_sheet(DETAIL_SHEET_NAME)

    rolling_total = summary["rolling_total"]
    wave_total = summary["wave_total"]
    rolling_types: dict[str, Decimal] = summary["rolling_types"]
    wave_types: dict[str, Decimal] = summary["wave_types"]

    rows = [
        ["排单分解", None, None, None],
        ["滚筒排单分解", None, "波轮排单分解", None],
        ["滚筒总数", _number_or_int(rolling_total), "波轮总数", _number_or_int(wave_total)],
    ]
    max_rows = max(len(ROLLING_LABELS), len(WAVE_LABELS))
    for index in range(max_rows):
        rolling_label = ROLLING_LABELS[index] if index < len(ROLLING_LABELS) else None
        wave_label = WAVE_LABELS[index] if index < len(WAVE_LABELS) else None
        rows.append(
            [
                rolling_label,
                _rolling_value_for_label(rolling_label, rolling_types),
                wave_label,
                _number_or_int(wave_types.get(wave_label, Decimal("0"))) if wave_label else None,
            ]
        )
    rows.append(["合计", f"=SUM(B5:B{3 + len(ROLLING_LABELS)})", "合计", f"=SUM(D4:D{3 + len(WAVE_LABELS)})"])

    for row in rows:
        sheet.append(row)

    sheet.merge_cells("A1:D1")
    sheet.merge_cells("A2:B2")
    sheet.merge_cells("C2:D2")
    _style_decomposition_sheet(sheet)
    _move_sheet_after(workbook, DETAIL_SHEET_NAME, main_sheet_name)
    logger.info(f"已创建“{DETAIL_SHEET_NAME}”，并写入滚筒/波轮总数与当前排单分解统计。")


def _apply_skd_style(sheet: Worksheet, row_index: int, headers: dict[str, int], type_name: str, is_dry: bool) -> None:
    _apply_decomposition_style(
        sheet,
        row_index,
        headers,
        type_name,
        SKD_MAIN_FILL,
        SKD_DRY_TYPE_FILL if is_dry else SKD_NORMAL_TYPE_FILL,
    )


def _apply_decomposition_style(
    sheet: Worksheet,
    row_index: int,
    headers: dict[str, int],
    type_name: str,
    main_fill_rgb: str,
    type_fill_rgb: str,
) -> None:
    main_fill = PatternFill(fill_type="solid", fgColor=main_fill_rgb)
    type_fill = PatternFill(fill_type="solid", fgColor=type_fill_rgb)
    for field in ["基本开始日期", "备注"]:
        cell = sheet.cell(row_index, headers[field])
        cell.fill = main_fill
        _set_font_for_fill(cell, main_fill_rgb)

    type_cell = sheet.cell(row_index, headers["类型"])
    type_cell.value = type_name
    type_cell.fill = type_fill
    _set_font_for_fill(type_cell, type_fill_rgb)


def _is_uncolored_for_decomposition(sheet: Worksheet, row_index: int, headers: dict[str, int]) -> bool:
    for field in DECOMPOSE_COLUMNS:
        cell = sheet.cell(row_index, headers[field])
        if cell.fill.fill_type is not None:
            return False
    return True


def _style_decomposition_sheet(sheet: Worksheet) -> None:
    blue_fill = PatternFill(fill_type="solid", fgColor="0070C0")
    section_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
    light_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=4):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if cell.row <= 2 else "left", vertical="center")
            cell.font = Font(color="000000", bold=False)

    sheet["A1"].fill = blue_fill
    sheet["A1"].font = Font(color="FFFFFF", bold=True)
    for cell in sheet[2]:
        cell.fill = section_fill
        cell.font = Font(bold=True, color="000000")

    shaded_labels = {
        "CKD含筒部装",
        "内销铁皮变频",
        "内销铁皮",
        "C6热泵干衣机",
        "复式烘干",
        "C6单洗",
        "C6Q10烘干",
        "普通烘干",
        "内销",
        "外销",
    }
    for row_index in range(3, sheet.max_row + 1):
        for label_col in (1, 3):
            value = sheet.cell(row_index, label_col).value
            if value in shaded_labels:
                sheet.cell(row_index, label_col).fill = light_fill
                sheet.cell(row_index, label_col + 1).fill = light_fill

    for cell in sheet[sheet.max_row]:
        cell.font = Font(bold=True, color="000000")
    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 14
    sheet.column_dimensions["C"].width = 22
    sheet.column_dimensions["D"].width = 14


def _collect_decomposition_summary(sheet: Worksheet, headers: dict[str, int]) -> dict[str, Any]:
    rolling_total = Decimal("0")
    wave_total = Decimal("0")
    rolling_types = _empty_type_totals(ROLLING_LABELS)
    wave_types = _empty_type_totals(WAVE_LABELS)

    for row_index in range(2, sheet.max_row + 1):
        line = str(sheet.cell(row_index, headers["线体"]).value or "").strip()
        order_qty = _to_decimal(sheet.cell(row_index, headers["订单数"]).value)
        type_name = str(sheet.cell(row_index, headers["类型"]).value or "").strip()

        if _is_wave_line(line):
            wave_total += order_qty
            if type_name:
                wave_types[type_name] = wave_types.get(type_name, Decimal("0")) + order_qty
        else:
            rolling_total += order_qty
            if type_name:
                rolling_types[type_name] = rolling_types.get(type_name, Decimal("0")) + order_qty

    return {
        "rolling_total": rolling_total,
        "wave_total": wave_total,
        "rolling_types": rolling_types,
        "wave_types": wave_types,
    }


def _empty_type_totals(labels: list[str]) -> dict[str, Decimal]:
    return {label: Decimal("0") for label in labels if label}


def _rolling_value_for_label(label: str | None, rolling_types: dict[str, Decimal]):
    if not label:
        return None
    if label == "SKD总数(包含烘干)":
        value = rolling_types.get("SKD", Decimal("0")) + rolling_types.get("SKD烘干", Decimal("0"))
    else:
        value = rolling_types.get(label, Decimal("0"))
    return _number_or_int(value)


def _product_kind_from_material_code(value: Any) -> str:
    material_code = str(value or "").strip().upper()
    if not material_code:
        return "unknown"
    marker_index = material_code.find("U")
    body = material_code[marker_index + 1 :] if marker_index >= 0 else material_code
    if body.startswith("605"):
        return "dryer"
    if body.startswith("60101"):
        return "wash"
    if body.startswith("60102"):
        return "dry"
    return "unknown"


def _is_t7p7t5p5_dreame_description(description: str) -> bool:
    text = str(description or "").strip()
    if not text:
        return False
    upper_text = text.upper()
    first_section = upper_text.split("/", 1)[0].strip()
    first_token = first_section.split()[0] if first_section.split() else first_section
    if any(series in first_token for series in ("T7", "P7", "T5", "P5")):
        return True
    return "追觅" in text or "DREAME" in upper_text or "DREMA" in upper_text


def _is_t9p9_description(description: str) -> bool:
    first_token = _description_first_token(description)
    return _contains_series_token(first_token, ("T9", "P9"))


def _description_first_token(description: str) -> str:
    text = str(description or "").strip().upper()
    if not text:
        return ""
    first_section = text.split("/", 1)[0].strip()
    return first_section.split()[0] if first_section.split() else first_section


def _contains_series_token(token: str, series_values: tuple[str, ...]) -> bool:
    for series in series_values:
        start = 0
        while True:
            index = token.find(series, start)
            if index < 0:
                break
            next_index = index + len(series)
            if next_index >= len(token) or not token[next_index].isdigit():
                return True
            start = index + 1
    return False


def _set_font_for_fill(cell, fill_rgb: str) -> None:
    font = copy(cell.font)
    font.color = "FFFFFF" if _is_dark_color(fill_rgb) else "000000"
    cell.font = font


def _is_dark_color(rgb: str) -> bool:
    value = rgb[-6:]
    r = int(value[0:2], 16)
    g = int(value[2:4], 16)
    b = int(value[4:6], 16)
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return luminance < 140


def _header_map(sheet: Worksheet) -> dict[str, int]:
    raw_headers = [normalize_header(cell.value) for cell in sheet[1]]
    headers = deduplicate_headers(raw_headers)
    return {header: index + 1 for index, header in enumerate(headers)}


def _require_columns(header_map: dict[str, int], columns: list[str], context: str) -> None:
    missing = [column for column in columns if column not in header_map]
    if missing:
        raise MissingRequiredFieldsError(f"{context}缺少关键字段：" + "、".join(missing))


def _is_wave_line(line: str) -> bool:
    return line in WAVE_LINES


def _to_decimal(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _number_or_int(value: Decimal):
    return int(value) if value == value.to_integral_value() else float(value)


def _move_sheet_after(workbook: Workbook, sheet_name: str, anchor_sheet_name: str) -> None:
    sheet = workbook[sheet_name]
    remaining = [item for item in workbook._sheets if item is not sheet]
    anchor_index = next(index for index, item in enumerate(remaining) if item.title == anchor_sheet_name)
    workbook._sheets = remaining[: anchor_index + 1] + [sheet] + remaining[anchor_index + 1 :]
