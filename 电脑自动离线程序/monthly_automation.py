# -*- coding: utf-8 -*-
"""Shared monthly automation helpers for a self-contained launcher folder."""

from __future__ import annotations

import json
import sys
import traceback
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Any


PROGRAM_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PROGRAM_DIR.parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from dingbian_classifier.pipeline import run


DEFAULT_CONFIG_PATH = PROGRAM_DIR / "monthly_flow.json"
DEFAULT_LOG_DIR = PROGRAM_DIR / "logs"
DEFAULT_INPUT_DIR = PROJECT_DIR / "inputs"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "outputs"
DEFAULT_STAGES = [
    "prepare-standard-units",
    "format-main-sheet",
    "decompose-skd",
    "decompose-rolling-remarks",
    "decompose-t7p7t5p5-dreame",
    "decompose-t9p9",
    "decompose-t9p9-dryer",
    "decompose-t10p10",
    "decompose-c6-heat-pump-dryer",
    "decompose-composite-penguin-c6",
    "decompose-rolling-final",
    "decompose-wave-basic",
    "decompose-wave-final",
    "decompose-extra-summary",
    "classify",
]
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


def load_monthly_flow(
    config_path: Path = DEFAULT_CONFIG_PATH,
    *,
    base_dir: Path = PROJECT_DIR,
) -> tuple[list[str], Path]:
    if not config_path.exists():
        return DEFAULT_STAGES.copy(), DEFAULT_OUTPUT_DIR

    with config_path.open("r", encoding="utf-8") as config_file:
        config = json.load(config_file)

    stages = _read_stages(config)
    output_dir = _read_output_dir(config, base_dir)
    return stages, output_dir


def _read_stages(config: dict[str, Any]) -> list[str]:
    raw_stages = config.get("stages", DEFAULT_STAGES)
    if not isinstance(raw_stages, list) or not raw_stages:
        raise ValueError("monthly_flow.json 中的 stages 必须是非空数组。")

    stages: list[str] = []
    for stage in raw_stages:
        if not isinstance(stage, str) or not stage.strip():
            raise ValueError("monthly_flow.json 中的每个 stage 都必须是非空文本。")
        stages.append(stage.strip())
    return stages


def _read_output_dir(config: dict[str, Any], base_dir: Path) -> Path:
    raw_output_dir = config.get("output_dir", str(DEFAULT_OUTPUT_DIR))
    if not isinstance(raw_output_dir, str) or not raw_output_dir.strip():
        raise ValueError("monthly_flow.json 中的 output_dir 必须是非空文本。")

    output_dir = Path(raw_output_dir.strip())
    if not output_dir.is_absolute():
        output_dir = base_dir / output_dir
    return output_dir


def ensure_input_dir(input_dir: Path = DEFAULT_INPUT_DIR) -> Path:
    input_dir.mkdir(exist_ok=True)
    return input_dir


def create_run_log_path(log_dir: Path = DEFAULT_LOG_DIR, prefix: str = "monthly_run") -> Path:
    log_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return log_dir / f"{prefix}_{timestamp}.log"


def write_log(log_path: Path, message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_path.parent.mkdir(exist_ok=True)
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")


def write_exception_log(log_path: Path, exc: BaseException) -> None:
    log_path.parent.mkdir(exist_ok=True)
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write("\n[详细错误]\n")
        log_file.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        log_file.write("\n")


def validate_input_file(input_path: Path) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    if input_path.name.startswith("~$"):
        raise ValueError("不能处理 Excel 临时锁文件（文件名以 ~$ 开头）。请关闭 Excel 后选择原始文件。")
    if input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        allowed = "、".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"请选择月度排产主 Excel 文件（支持：{allowed}）。")


def run_stage(
    input_path: Path,
    output_dir: Path,
    stage: str,
    log_path: Path,
    *,
    coefficient_lookup: Path | None = None,
    sheet_metal_lookup: Path | None = None,
    sheet_metal_bom_lookup: Path | None = None,
    material_description_lookup: Path | None = None,
) -> Path:
    validate_input_file(input_path)
    output_dir.mkdir(exist_ok=True)
    write_log(log_path, f"开始阶段：{stage}")
    write_log(log_path, f"阶段输入：{input_path}")
    if coefficient_lookup is not None:
        write_log(log_path, f"系数查询表：{coefficient_lookup}")
    if sheet_metal_lookup is not None:
        write_log(log_path, f"钣金型号查询表：{sheet_metal_lookup}")
    if sheet_metal_bom_lookup is not None:
        write_log(log_path, f"钣金型号BOM表：{sheet_metal_bom_lookup}")
    if material_description_lookup is not None:
        write_log(log_path, f"物料描述查询表：{material_description_lookup}")
    try:
        with log_path.open("a", encoding="utf-8") as log_file:
            with redirect_stdout(log_file), redirect_stderr(log_file):
                output_path = run(
                    input_path,
                    output_dir,
                    stage=stage,
                    coefficient_lookup=coefficient_lookup,
                    sheet_metal_lookup=sheet_metal_lookup,
                    sheet_metal_bom_lookup=sheet_metal_bom_lookup,
                    material_description_lookup=material_description_lookup,
                )
    except Exception as exc:
        write_log(log_path, f"阶段失败：{stage}；原因：{exc}")
        raise
    write_log(log_path, f"阶段完成：{stage}")
    write_log(log_path, f"阶段输出：{output_path}")
    return output_path


def run_monthly_flow(input_path: Path, stages: list[str], output_dir: Path, log_path: Path) -> Path:
    current_input = input_path
    latest_output = input_path
    for stage in stages:
        latest_output = run_stage(current_input, output_dir, stage, log_path)
        current_input = latest_output
    return latest_output
