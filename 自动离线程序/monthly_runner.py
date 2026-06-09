# -*- coding: utf-8 -*-
"""On-demand monthly Excel automation launcher."""

from __future__ import annotations

import json
import os
import sys
import traceback
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox

LAUNCHER_DIR = Path(__file__).resolve().parent
PROJECT_DIR = LAUNCHER_DIR.parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from dingbian_classifier.exceptions import ClassifierError
from dingbian_classifier.pipeline import run


CONFIG_PATH = LAUNCHER_DIR / "monthly_flow.json"
LOG_DIR = LAUNCHER_DIR / "logs"
DEFAULT_INPUT_DIR = PROJECT_DIR / "inputs"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "outputs"
DEFAULT_STAGES = ["decompose-extra-summary", "classify"]
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def load_monthly_flow() -> tuple[list[str], Path]:
    if not CONFIG_PATH.exists():
        return DEFAULT_STAGES.copy(), DEFAULT_OUTPUT_DIR

    with CONFIG_PATH.open("r", encoding="utf-8") as config_file:
        config = json.load(config_file)

    stages = _read_stages(config)
    output_dir = _read_output_dir(config)
    return stages, output_dir


def create_run_log_path() -> Path:
    LOG_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return LOG_DIR / f"monthly_run_{timestamp}.log"


def write_log(log_path: Path, message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")


def write_exception_log(log_path: Path, exc: BaseException) -> None:
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write("\n[详细错误]\n")
        log_file.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        log_file.write("\n")


def ensure_input_dir() -> Path:
    DEFAULT_INPUT_DIR.mkdir(exist_ok=True)
    return DEFAULT_INPUT_DIR


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


def _read_output_dir(config: dict[str, Any]) -> Path:
    raw_output_dir = config.get("output_dir", str(DEFAULT_OUTPUT_DIR))
    if not isinstance(raw_output_dir, str) or not raw_output_dir.strip():
        raise ValueError("monthly_flow.json 中的 output_dir 必须是非空文本。")

    output_dir = Path(raw_output_dir.strip())
    if not output_dir.is_absolute():
        output_dir = PROJECT_DIR / output_dir
    return output_dir


def choose_input_file() -> Path | None:
    selected = filedialog.askopenfilename(
        title="请选择本次月度定编 Excel 文件",
        initialdir=str(DEFAULT_INPUT_DIR),
        filetypes=[
            ("Excel 文件", "*.xlsx *.xlsm *.xls"),
            ("所有文件", "*.*"),
        ],
    )
    if not selected:
        return None
    return Path(selected)


def validate_input_file(input_path: Path) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    if input_path.name.startswith("~$"):
        raise ValueError("不能处理 Excel 临时锁文件（文件名以 ~$ 开头）。请关闭 Excel 后选择原始文件。")
    if input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        allowed = "、".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"请选择 Excel 文件（支持：{allowed}）。")


def run_monthly_flow(input_path: Path, stages: list[str], output_dir: Path, log_path: Path) -> Path:
    output_dir.mkdir(exist_ok=True)
    current_input = input_path
    latest_output = input_path
    for index, stage in enumerate(stages, start=1):
        write_log(log_path, f"开始阶段 {index}/{len(stages)}：{stage}")
        write_log(log_path, f"阶段输入：{current_input}")
        try:
            with log_path.open("a", encoding="utf-8") as log_file:
                with redirect_stdout(log_file), redirect_stderr(log_file):
                    latest_output = run(current_input, output_dir, stage=stage)
        except Exception as exc:
            write_log(log_path, f"阶段失败：{stage}；原因：{exc}")
            raise
        write_log(log_path, f"阶段完成：{stage}")
        write_log(log_path, f"阶段输出：{latest_output}")
        current_input = latest_output
    return latest_output


def open_output_dir(output_dir: Path, log_path: Path) -> None:
    output_dir.mkdir(exist_ok=True)
    try:
        os.startfile(str(output_dir))
        write_log(log_path, f"已打开输出目录：{output_dir}")
    except OSError as exc:
        write_log(log_path, f"打开输出目录失败：{exc}")


def show_error(title: str, exc: BaseException) -> None:
    detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
    messagebox.showerror(title, detail or str(exc))


def main() -> int:
    log_path = create_run_log_path()
    write_log(log_path, "月度离线处理启动")
    write_log(log_path, f"项目目录：{PROJECT_DIR}")
    write_log(log_path, f"启动目录：{LAUNCHER_DIR}")

    root = tk.Tk()
    root.withdraw()

    try:
        ensure_input_dir()
        stages, output_dir = load_monthly_flow()
        write_log(log_path, f"流程配置：{' -> '.join(stages)}")
        write_log(log_path, f"输出目录：{output_dir}")
        input_path = choose_input_file()
        if input_path is None:
            write_log(log_path, "用户取消选择文件，程序退出。")
            root.destroy()
            return 0

        write_log(log_path, f"用户选择文件：{input_path}")
        validate_input_file(input_path)
        output_path = run_monthly_flow(input_path, stages, output_dir, log_path)
    except (ClassifierError, FileNotFoundError, ValueError, json.JSONDecodeError) as exc:
        write_log(log_path, f"处理失败：{exc}")
        write_exception_log(log_path, exc)
        show_error("月度处理失败", exc)
        root.destroy()
        return 1
    except Exception as exc:
        write_log(log_path, f"未知错误：{exc}")
        write_exception_log(log_path, exc)
        show_error("月度处理发生未知错误", exc)
        root.destroy()
        return 1

    stage_text = " -> ".join(stages)
    write_log(log_path, f"月度离线处理完成：{output_path}")
    open_output_dir(output_dir, log_path)
    messagebox.showinfo(
        "月度处理完成",
        f"处理流程：{stage_text}\n\n输出文件：\n{output_path}\n\n日志文件：\n{log_path}",
    )
    root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
