# -*- coding: utf-8 -*-
"""Streamlit web entry for the mobile/cloud monthly Excel automation."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import streamlit as st

from monthly_automation import (
    DEFAULT_CONFIG_PATH,
    load_monthly_flow,
    run_monthly_flow,
    validate_input_file,
    write_exception_log,
    write_log,
)


APP_TITLE = "月度 Excel 自动处理"
PASSWORD_ENV = "DINGBIAN_APP_PASSWORD"


def require_password() -> bool:
    expected = os.environ.get(PASSWORD_ENV) or st.secrets.get("app_password", "")
    if not expected:
        return True

    password = st.text_input("访问密码", type="password")
    if password == expected:
        return True
    if password:
        st.error("访问密码不正确。")
    return False


def save_upload(uploaded_file, target_dir: Path) -> Path:
    target_dir.mkdir(exist_ok=True)
    input_path = target_dir / uploaded_file.name
    input_path.write_bytes(uploaded_file.getbuffer())
    return input_path


def process_upload(uploaded_file) -> tuple[str, bytes, str, bytes]:
    work_dir = Path(tempfile.mkdtemp(prefix="dingbian_monthly_"))
    input_dir = work_dir / "inputs"
    output_dir = work_dir / "outputs"
    log_path = work_dir / "logs" / "web_run.log"

    input_path = save_upload(uploaded_file, input_dir)
    validate_input_file(input_path)

    stages, _configured_output_dir = load_monthly_flow(DEFAULT_CONFIG_PATH)
    write_log(log_path, "云端网页处理启动")
    write_log(log_path, f"上传文件：{uploaded_file.name}")
    write_log(log_path, f"流程配置：{' -> '.join(stages)}")
    output_path = run_monthly_flow(input_path, stages, output_dir, log_path)

    return (
        output_path.name,
        output_path.read_bytes(),
        f"{output_path.stem}_web_run.log",
        log_path.read_bytes(),
    )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="centered")
    st.title(APP_TITLE)

    if not require_password():
        st.stop()

    stages, _output_dir = load_monthly_flow(DEFAULT_CONFIG_PATH)
    st.caption("当前流程：" + " -> ".join(stages))

    uploaded_file = st.file_uploader(
        "上传基础数据已补齐、可进入完整分解的 Excel 当前结果文件",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False,
    )

    if uploaded_file is None:
        st.info("请上传 Excel 文件后开始处理。")
        return

    if uploaded_file.name.startswith("~$"):
        st.error("不能处理 Excel 临时锁文件（文件名以 ~$ 开头）。")
        return

    if st.button("开始处理", type="primary"):
        with st.spinner("正在处理，请保持页面打开..."):
            try:
                output_name, output_bytes, log_name, log_bytes = process_upload(uploaded_file)
            except Exception as exc:
                st.error(f"处理失败：{exc}")
                error_log_path = Path(tempfile.gettempdir()) / "dingbian_web_error.log"
                error_log_path.write_text(f"处理失败：{exc}\n", encoding="utf-8")
                write_exception_log(error_log_path, exc)
                st.download_button(
                    "下载错误日志",
                    data=error_log_path.read_bytes(),
                    file_name=error_log_path.name,
                    mime="text/plain",
                )
                return

        st.success("处理完成。")
        st.download_button(
            "下载处理结果",
            data=output_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "下载运行日志",
            data=log_bytes,
            file_name=log_name,
            mime="text/plain",
        )


if __name__ == "__main__":
    main()
