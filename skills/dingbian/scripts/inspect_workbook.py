#!/usr/bin/env python
"""Read-only inspection for Dingbian production-plan workbooks."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import openpyxl


TARGET_PATTERN = re.compile(r"^W\d+\s*-\s*W\d+.*周排产明细")
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
PENDING_GROUPS = {
    "coefficient_pending": ("系数仍缺失", "系数补充"),
    "sheet_metal_pending": ("钣金型号补充",),
    "material_description_pending": ("物料描述仍缺失", "物料描述补充"),
}
FINAL_SHEETS = ("排单分解表明细", "各线体分类明细表", "分类结果汇总表")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="只读检查定编工作簿状态并输出 JSON")
    parser.add_argument("workbook", type=Path, help="待检查的 .xlsx 或 .xlsm 文件")
    return parser.parse_args()


def fail(message: str, exit_code: int = 2) -> int:
    print(json.dumps({"ok": False, "error": message}, ensure_ascii=False, indent=2))
    return exit_code


def count_data_rows(sheet: Any) -> int:
    count = 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if any(value is not None and str(value).strip() for value in row):
            count += 1
    return count


def pending_count(workbook: Any, sheet_names: tuple[str, ...]) -> int:
    return max(
        (count_data_rows(workbook[name]) for name in sheet_names if name in workbook.sheetnames),
        default=0,
    )


def main() -> int:
    args = parse_args()
    raw_path = args.workbook
    if raw_path.name.startswith("~$"):
        return fail("不能检查 Excel 临时锁文件；请关闭 Excel 后使用原始文件。")
    path = raw_path.expanduser().resolve()

    if not path.exists():
        return fail(f"文件不存在：{path}")
    if path.suffix.lower() not in SUPPORTED_SUFFIXES:
        return fail("主工作簿仅支持 .xlsx 或 .xlsm。")

    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        return fail(f"无法打开工作簿：{type(exc).__name__}: {exc}")

    try:
        target_sheets = [name for name in workbook.sheetnames if TARGET_PATTERN.search(name)]
        result: dict[str, Any] = {
            "ok": len(target_sheets) == 1,
            "path": str(path),
            "target_sheets": target_sheets,
            "target_sheet": target_sheets[0] if len(target_sheets) == 1 else None,
            "target_data_rows": (
                count_data_rows(workbook[target_sheets[0]]) if len(target_sheets) == 1 else None
            ),
            "sheet_count": len(workbook.sheetnames),
        }

        for key, sheet_names in PENDING_GROUPS.items():
            result[key] = pending_count(workbook, sheet_names)

        result["unclassified_rows"] = (
            count_data_rows(workbook["未分类数据"]) if "未分类数据" in workbook.sheetnames else 0
        )
        result["final_sheets"] = {
            name: name in workbook.sheetnames for name in FINAL_SHEETS
        }

        if not target_sheets:
            result["error"] = "未找到符合命名规则的周排产明细主工作表。"
            result["sheet_names"] = workbook.sheetnames
        elif len(target_sheets) > 1:
            result["error"] = "找到多个候选主工作表，需要用户指定。"

        if result["coefficient_pending"] > 0:
            result["next_action"] = "resolve_coefficients"
        elif result["sheet_metal_pending"] > 0:
            result["next_action"] = "resolve_sheet_metal"
        elif result["material_description_pending"] > 0:
            result["next_action"] = "resolve_material_description"
        elif all(result["final_sheets"].values()):
            result["next_action"] = "complete"
        else:
            result["next_action"] = "continue_processing"

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["ok"] else 3
    finally:
        workbook.close()


if __name__ == "__main__":
    sys.exit(main())
