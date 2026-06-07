"""Header and row normalization helpers."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any

from .exceptions import MissingRequiredFieldsError

REQUIRED_FIELDS = ["物料编码", "物料描述", "钣金型号", "线体", "订单数"]


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def deduplicate_headers(headers: list[str]) -> list[str]:
    """Make duplicate/blank headers usable while keeping original order."""
    seen: dict[str, int] = {}
    result: list[str] = []
    for index, header in enumerate(headers, start=1):
        base = header or f"空列{index}"
        count = seen.get(base, 0) + 1
        seen[base] = count
        result.append(base if count == 1 else f"{base}_{count}")
    return result


def validate_required_fields(headers: list[str]) -> None:
    available = set(headers)
    missing = [field for field in REQUIRED_FIELDS if field not in available]
    if missing:
        raise MissingRequiredFieldsError("主数据表缺少关键字段：" + "、".join(missing))


def to_number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(Decimal(str(value).replace(",", "").strip()))
    except (InvalidOperation, ValueError):
        return 0.0


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()

