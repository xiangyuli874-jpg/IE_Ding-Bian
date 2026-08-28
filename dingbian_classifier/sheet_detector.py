"""Target worksheet detection."""

from __future__ import annotations

import re

from .exceptions import TargetSheetNotFoundError
from .logger import ProcessingLogger

TARGET_SHEET_PATTERN = re.compile(r"^W\d+\s*-\s*W\d+.*周排产明细")


def find_target_sheet(
    sheet_names: list[str],
    logger: ProcessingLogger,
    target_sheet_name: str | None = None,
) -> str:
    """Return the first sheet name matching the production detail naming rule."""
    if target_sheet_name:
        if target_sheet_name not in sheet_names:
            raise TargetSheetNotFoundError(f"指定的主工作表不存在：{target_sheet_name}")
        logger.info(f"使用指定目标工作表：{target_sheet_name}")
        return target_sheet_name

    matches = [name for name in sheet_names if TARGET_SHEET_PATTERN.search(name)]
    if not matches:
        raise TargetSheetNotFoundError(
            "未找到符合命名规则的工作表：W数字-W数字...周排产明细"
        )

    if len(matches) > 1:
        logger.warning(
            "找到多个符合条件的周排产明细工作表，默认处理第一个："
            + "、".join(matches)
        )
    else:
        logger.info(f"找到目标工作表：{matches[0]}")

    return matches[0]
