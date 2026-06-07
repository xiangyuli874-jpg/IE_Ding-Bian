"""Classification engine."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .data_cleaner import normalize_text, to_number
from .rule_config import ClassificationRule


@dataclass
class ClassificationResult:
    categories: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    output_names: dict[str, str] = field(default_factory=dict)
    unmatched: list[dict[str, Any]] = field(default_factory=list)


def row_matches_rule(row: dict[str, Any], rule: ClassificationRule) -> bool:
    text = normalize_text(row.get(rule.field_name))
    if rule.include_keywords and not any(keyword in text for keyword in rule.include_keywords):
        return False
    if rule.exclude_keywords and any(keyword in text for keyword in rule.exclude_keywords):
        return False
    return bool(rule.include_keywords or text)


def deduplicate_rows(rows: list[dict[str, Any]], fields: list[str]) -> list[dict[str, Any]]:
    if not fields:
        return rows
    seen: set[tuple[Any, ...]] = set()
    result: list[dict[str, Any]] = []
    for row in rows:
        key = tuple(row.get(field) for field in fields)
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def summarize_rows(rows: list[dict[str, Any]], dimensions: list[str]) -> list[dict[str, Any]]:
    if not dimensions:
        dimensions = ["线体", "周次"]
    buckets: dict[tuple[Any, ...], float] = {}
    for row in rows:
        key = tuple(row.get(field) for field in dimensions)
        buckets[key] = buckets.get(key, 0.0) + to_number(row.get("订单数"))

    result: list[dict[str, Any]] = []
    for key, qty in buckets.items():
        item = {dimensions[index]: key[index] for index in range(len(dimensions))}
        item["订单数合计"] = qty
        result.append(item)
    return result


def classify_rows(rows: list[dict[str, Any]], rules: list[ClassificationRule]) -> ClassificationResult:
    result = ClassificationResult()
    for row in rows:
        matched_rule: ClassificationRule | None = None
        for rule in rules:
            if row_matches_rule(row, rule):
                matched_rule = rule
                break

        if matched_rule is None:
            result.unmatched.append(row)
            continue

        bucket = result.categories.setdefault(matched_rule.category_name, [])
        bucket.append(row)
        result.output_names[matched_rule.category_name] = matched_rule.output_sheet_name

    for rule in rules:
        rows_for_rule = result.categories.get(rule.category_name)
        if not rows_for_rule:
            continue
        if rule.deduplicate:
            rows_for_rule = deduplicate_rows(rows_for_rule, rule.deduplicate_fields)
        if rule.summarize_order_qty:
            rows_for_rule = summarize_rows(rows_for_rule, rule.summary_fields)
        result.categories[rule.category_name] = rows_for_rule

    return result

