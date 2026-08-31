"""Optional Excel compatibility pass using Microsoft Excel COM on Windows."""

from __future__ import annotations

import os
import platform
import subprocess
from pathlib import Path

from .logger import ProcessingLogger


def ps_quote(path: Path) -> str:
    return "'" + str(path).replace("'", "''") + "'"


def resave_with_excel_if_available(path: Path, logger: ProcessingLogger) -> None:
    """Ask desktop Excel to resave the workbook so Excel-specific records stay valid."""
    if os.environ.get("DINGBIAN_SKIP_EXCEL_RESAVE") == "1":
        logger.info("已跳过中间阶段的 Excel 兼容性另存。")
        return
    if platform.system() != "Windows":
        return

    temp_path = path.with_name(f"{path.stem}_excel_resave{path.suffix}")
    script = f"""
$src = {ps_quote(path)}
$dst = {ps_quote(temp_path)}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
  $wb = $excel.Workbooks.Open($src, 0, $false, 5, "", "", $true, 1, "", $false, $false, 0, $false, $true, 1)
  $mainSheet = $wb.ActiveSheet
  $mainSheet.Activate()
  $excel.ActiveWindow.FreezePanes = $false
  $excel.ActiveWindow.Split = $false
  $excel.ActiveWindow.SplitColumn = 0
  $excel.ActiveWindow.SplitRow = 0
  $mainSheet.Range("A2").Select()
  $excel.ActiveWindow.SplitRow = 1
  $excel.ActiveWindow.FreezePanes = $true
  $wb.SaveAs($dst, 51)
  $wb.Close($false)
}} finally {{
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
        check=False,
    )
    if completed.returncode != 0 or not temp_path.exists():
        detail = (completed.stderr or completed.stdout or "").strip().splitlines()
        suffix = f" 原因：{detail[-1]}" if detail else ""
        logger.warning("Excel 兼容性另存未完成，保留 openpyxl 生成的结果文件。" + suffix)
        return

    os.replace(temp_path, path)
    logger.info("已使用 Microsoft Excel 重新另存结果文件，提高打开兼容性。")
