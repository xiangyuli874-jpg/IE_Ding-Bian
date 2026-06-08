"""Output backup and retention helpers."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from .logger import ProcessingLogger

DEFAULT_MAX_BACKUPS = 3
HISTORY_DIR_NAME = "history"


def backup_current_result(
    input_path: Path,
    output_dir: Path,
    stage: str,
    logger: ProcessingLogger,
    max_backups: int = DEFAULT_MAX_BACKUPS,
) -> Path | None:
    """Back up the current result workbook before creating a new output."""
    source = choose_backup_source(input_path, output_dir)
    if source is None:
        logger.info("未找到需要备份的当前结果文件。")
        return None

    history_dir = output_dir / HISTORY_DIR_NAME
    history_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = history_dir / f"{source.stem}_before_{stage}_{timestamp}{source.suffix}"
    shutil.copy2(source, backup_path)
    logger.info(f"已备份当前结果文件：{backup_path}")
    prune_history(history_dir, max_backups, logger)
    return backup_path


def choose_backup_source(input_path: Path, output_dir: Path) -> Path | None:
    """Prefer the input file if it is an output result; otherwise use latest output."""
    if input_path.parent.resolve() == output_dir.resolve() and is_result_workbook(input_path):
        return input_path if input_path.exists() else None

    candidates = result_workbooks(output_dir)
    return candidates[-1] if candidates else None


def cleanup_output_results(output_dir: Path, keep_path: Path, logger: ProcessingLogger) -> None:
    """Keep only the latest/current result workbook in the output directory."""
    for path in result_workbooks(output_dir):
        if path.resolve() == keep_path.resolve():
            continue
        try:
            path.unlink()
            logger.info(f"已清理旧结果副本：{path}")
        except PermissionError:
            logger.warning(f"旧结果副本正在被占用，暂未删除：{path}")


def prune_history(history_dir: Path, max_backups: int, logger: ProcessingLogger) -> None:
    backups = sorted(
        [path for path in history_dir.glob("*.xlsx") if path.is_file()],
        key=lambda path: path.stat().st_mtime,
    )
    while len(backups) > max_backups:
        oldest = backups.pop(0)
        try:
            oldest.unlink()
            logger.info(f"已删除最早的历史备份：{oldest}")
        except PermissionError:
            logger.warning(f"历史备份正在被占用，暂未删除：{oldest}")


def result_workbooks(output_dir: Path) -> list[Path]:
    return sorted(
        [
            path
            for path in output_dir.glob("*.xlsx")
            if path.is_file() and not path.name.startswith("~$") and is_result_workbook(path)
        ],
        key=lambda path: path.stat().st_mtime,
    )


def is_result_workbook(path: Path) -> bool:
    return "_分类结果_" in path.stem
