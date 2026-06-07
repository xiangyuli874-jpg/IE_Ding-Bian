"""Target worksheet detection."""

from __future__ import annotations

import re

from .exceptions import TargetSheetNotFoundError
from .logger import ProcessingLogger

TARGET_SHEET_PATTERN = re.compile(r"^W\d+\s*-\s*W\d+.*周排产明细")


def find_target_sheet(sheet_names: list[str], logger: ProcessingLogger) -> str:
    """Return the first sheet name matching the production detail naming rule."""
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

