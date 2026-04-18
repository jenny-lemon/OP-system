import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

from performance_report import (
    generate_sales_report,
    load_daily_history_for_current_month,
    delete_daily_history_rows,
    load_output_file_log,
    LATEST_DIR,
)

TZ = timezone(timedelta(hours=8))


def now():
    return datetime.now(TZ)


def run_cmd(cmd):
    p = subprocess.run(cmd, shell=True, text=True, capture_output=True)
    return p.returncode, p.stdout, p.stderr


# ─────────────────────────────────────────────
# 主控表
# ─────────────────────────────────────────────
def render_main_page():
    st.title("📊 主控表")

    tasks = [
        ("業績報表", f"{sys.executable} performance_report.py dashboard false"),
    ]

    if "last_result" not in st.session_state:
        st.session_state.last_result = None

    for name, cmd in tasks:
        if st.button(f"▶ 執行 {name}"):
            with st.spinner("執行中..."):
                code, out, err = run_cmd(cmd)

            st.session_state.last_result = {
                "name": name,
                "code": code,
                "stdout": out,
                "stderr": err,
                "time": now().strftime("%H:%M:%S"),
            }

            st.rerun()

    # 統一顯示結果
    r = st.session_state.last_result
    if r:
        st.markdown("### 🖥 最近一次執行")
        st.write(r["name"], r["time"])
        st.code(r["stdout"] or r["stderr"] or "(無輸出)")


# ─────────────────────────────────────────────
# 業績報表
# ─────────────────────────────────────────────
def render_sales_page():
    st.title("📈 業績報表")

    c1, c2 = st.columns(2)

    with c1:
        if st.button("🔄 更新資料"):
            with st.spinner("更新中..."):
                result = generate_sales_report(
                    send_email=False,
                    persist_dashboard=True,
                    trigger="dashboard",
                )
            st.session_state.sales = result

    with c2:
        if st.button("📂 重新載入"):
            st.rerun()

    result = st.session_state.get("sales")

    if not result:
        st.info("請先更新資料")
        return

    df4 = result["df4"]
    daily_df = result["daily_df"]

    st.subheader("📊 月度摘要")
    st.dataframe(df4, use_container_width=True)

    st.subheader("📅 當日業績")
    st.dataframe(daily_df, use_container_width=True)

    # 🔥 改這裡：使用 daily_history
    st.subheader("📅 當月每日業績總覽（累積）")

    history_df = load_daily_history_for_current_month()

    if history_df.empty:
        st.info("目前沒有紀錄")
    else:
        ids = history_df["id"].astype(str).tolist()

        selected = st.multiselect("勾選刪除", ids)

        if st.button("🗑 刪除"):
            deleted = delete_daily_history_rows(selected)
            st.success(f"刪除 {deleted} 筆")
            st.rerun()

        st.dataframe(history_df, use_container_width=True)


# ─────────────────────────────────────────────
# 輸出檔案
# ─────────────────────────────────────────────
def render_output_page():
    st.title("📂 輸出檔案")

    st.subheader("📄 最新檔案")

    latest_dir = Path(LATEST_DIR)
    files = list(latest_dir.glob("*"))

    rows = []
    for f in files:
        rows.append({
            "檔名": f.name,
            "大小": f.stat().st_size,
            "時間": datetime.fromtimestamp(f.stat().st_mtime).strftime("%m/%d %H:%M"),
            "路徑": str(f),
        })

    st.dataframe(pd.DataFrame(rows), use_container_width=True)

    # 🔥 新功能
    st.subheader("🧾 輸出檔案紀錄")

    log_df = load_output_file_log()

    if log_df.empty:
        st.info("尚無紀錄")
    else:
        keyword = st.text_input("搜尋")

        if keyword:
            log_df = log_df[
                log_df["檔名"].str.contains(keyword, case=False, na=False)
                | log_df["完整路徑"].str.contains(keyword, case=False, na=False)
            ]

        st.dataframe(log_df, use_container_width=True)


# ─────────────────────────────────────────────
# Router
# ─────────────────────────────────────────────
def render_page(page):
    if page == "主控表":
        render_main_page()
    elif page == "業績報表":
        render_sales_page()
    elif page == "輸出檔案":
        render_output_page()
    else:
        render_main_page()
