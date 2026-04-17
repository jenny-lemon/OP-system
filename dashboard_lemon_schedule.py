import json
import math
import os
from datetime import datetime

import pandas as pd
import streamlit as st

from 業績報表 import (
    generate_sales_report,
    load_execution_log_for_current_month,
    load_daily_history_for_current_month,
    delete_daily_history_rows,
    LATEST_DIR,
)


# =========================
# 基本設定
# =========================
st.set_page_config(page_title="業績報表", page_icon="📊", layout="wide")

TITLE = "業績報表"
PAGE_SIZE = 10


# =========================
# 工具函式
# =========================
def file_exists(path):
    return os.path.exists(path)


def load_latest_payload():
    df4_path = os.path.join(LATEST_DIR, "df4.csv")
    daily_path = os.path.join(LATEST_DIR, "daily_df.csv")
    meta_path = os.path.join(LATEST_DIR, "meta.json")
    html_path = os.path.join(LATEST_DIR, "email_preview.html")

    payload = {
        "df4": pd.DataFrame(),
        "daily_df": pd.DataFrame(),
        "meta": {},
        "email_html": "",
    }

    if file_exists(df4_path):
        payload["df4"] = pd.read_csv(df4_path, encoding="utf-8-sig")

    if file_exists(daily_path):
        payload["daily_df"] = pd.read_csv(daily_path, encoding="utf-8-sig")

    if file_exists(meta_path):
        with open(meta_path, "r", encoding="utf-8") as f:
            payload["meta"] = json.load(f)

    if file_exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            payload["email_html"] = f.read()

    return payload


def fmt_int(x):
    try:
        return f"{int(float(x)):,}"
    except Exception:
        return "0"


def fmt_pct(x):
    try:
        return f"{float(x):.2%}"
    except Exception:
        return "0.00%"


def style_df4_display(df):
    if df.empty:
        return df

    out = df.copy()

    for col in ["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"]:
        if col in out.columns:
            out[col] = out[col].apply(fmt_int)

    for col in ["本月佔比", "次月佔比"]:
        if col in out.columns:
            out[col] = out[col].apply(fmt_pct)

    return out


def style_daily_display(df):
    if df.empty:
        return df

    out = df.copy()

    for col in out.columns:
        if col.endswith("業績") or col == "全區合計":
            out[col] = out[col].apply(fmt_int)
        elif col.endswith("佔比"):
            out[col] = out[col].apply(fmt_pct)

    return out


def style_exec_log_display(df):
    if df.empty:
        return df

    out = df.copy()

    for col in ["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"]:
        if col in out.columns:
            out[col] = out[col].apply(fmt_int)

    return out


def style_daily_history_display(df):
    if df.empty:
        return df

    out = df.copy()
    if "今日全區合計" in out.columns:
        out["今日全區合計"] = out["今日全區合計"].apply(fmt_int)

    return out


def get_total_row(df4):
    if df4.empty:
        return None

    total = df4[df4["城市"] == "加總"]
    if total.empty:
        return None

    return total.iloc[0]


def render_kpi_cards(df4):
    total = get_total_row(df4)

    if total is None:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("本月加總", "0")
        c2.metric("次月加總", "0")
        c3.metric("本月家電加總", "0")
        c4.metric("儲值金", "0")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("本月加總", fmt_int(total.get("本月加總", 0)))
    c2.metric("次月加總", fmt_int(total.get("次月加總", 0)))
    c3.metric("本月家電加總", fmt_int(total.get("本月家電加總", 0)))
    c4.metric("儲值金", fmt_int(total.get("儲值金", 0)))


def paginate_df(df, page_key, page_size=PAGE_SIZE):
    if df.empty:
        return df, 1, 1, 1

    total_rows = len(df)
    total_pages = max(1, math.ceil(total_rows / page_size))

    if page_key not in st.session_state:
        st.session_state[page_key] = 1

    st.session_state[page_key] = max(1, min(st.session_state[page_key], total_pages))
    page = st.session_state[page_key]

    start = (page - 1) * page_size
    end = start + page_size

    return df.iloc[start:end].copy(), page, total_pages, total_rows


def reset_pages():
    st.session_state["exec_page"] = 1
    st.session_state["daily_history_page"] = 1


# =========================
# 樣式
# =========================
st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1600px;
}
h1, h2, h3 {
    font-weight: 800 !important;
    color: #16243b;
}
.stButton > button {
    width: 100%;
    background-color: #16243b;
    color: white;
    border-radius: 10px;
    border: none;
    height: 54px;
    font-size: 20px;
    font-weight: 700;
}
.info-box {
    background: #dfe8f7;
    color: #1c5ea8;
    padding: 18px 22px;
    border-radius: 12px;
    font-size: 18px;
    font-weight: 700;
    margin-top: 6px;
    margin-bottom: 10px;
}
.small-note {
    color: #6b7280;
    font-size: 14px;
}
.section-gap {
    height: 10px;
}
</style>
""", unsafe_allow_html=True)


# =========================
# Session State
# =========================
if "exec_page" not in st.session_state:
    st.session_state["exec_page"] = 1

if "daily_history_page" not in st.session_state:
    st.session_state["daily_history_page"] = 1

if "delete_ids" not in st.session_state:
    st.session_state["delete_ids"] = []


# =========================
# 標題
# =========================
st.title("業績報表")
st.caption("LATEST + EXECUTION LOG")


# =========================
# 按鈕區
# =========================
b1, b2, b3 = st.columns(3)

with b1:
    if st.button("🔄 更新資料"):
        with st.spinner("更新資料中..."):
            result = generate_sales_report(
                send_email=False,
                persist_dashboard=True,
                trigger="dashboard",
            )
        reset_pages()
        st.success(f"更新完成：{result['updated_at']}")

with b2:
    if st.button("📧 寄送目前結果"):
        with st.spinner("寄送中..."):
            result = generate_sales_report(
                send_email=True,
                persist_dashboard=True,
                trigger="dashboard",
            )
        reset_pages()
        st.success(f"已寄送：{result['updated_at']}")

with b3:
    if st.button("📂 重新讀取已存資料"):
        reset_pages()
        st.success("已重新讀取")


# =========================
# 載入資料
# =========================
payload = load_latest_payload()
df4 = payload["df4"]
daily_df = payload["daily_df"]
meta = payload["meta"]
email_html = payload["email_html"]

execution_log_df = load_execution_log_for_current_month()
daily_history_df = load_daily_history_for_current_month()

updated_at = meta.get("updated_at", "尚未產生資料")

st.markdown(
    f'<div class="info-box">最新更新時間：{updated_at}</div>',
    unsafe_allow_html=True,
)


# =========================
# KPI
# =========================
render_kpi_cards(df4)

st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)


# =========================
# 各區月度摘要（信中的結果）
# =========================
st.header("各區月度摘要")

if df4.empty:
    st.warning("目前沒有各區月度摘要資料")
else:
    show_df4 = style_df4_display(df4)
    st.dataframe(show_df4, use_container_width=True, hide_index=True)


# =========================
# 當月每日業績總覽
# =========================
st.header("當月每日業績總覽")

if daily_df.empty:
    st.warning("目前沒有當月每日業績總覽資料")
else:
    show_daily = style_daily_display(daily_df)
    st.dataframe(show_daily, use_container_width=True, hide_index=True)


# =========================
# 當月累積執行紀錄
# =========================
st.header("當月累積執行紀錄")

exec_page_df, exec_page, exec_total_pages, exec_total_rows = paginate_df(
    execution_log_df,
    page_key="exec_page",
    page_size=10,
)

ec1, ec2, ec3 = st.columns([1, 1, 3])
with ec1:
    if st.button("◀ 上一頁", key="exec_prev"):
        st.session_state["exec_page"] = max(1, st.session_state["exec_page"] - 1)
        st.rerun()
with ec2:
    if st.button("下一頁 ▶", key="exec_next"):
        st.session_state["exec_page"] = min(exec_total_pages, st.session_state["exec_page"] + 1)
        st.rerun()
with ec3:
    st.markdown(
        f"<div style='font-size:18px;padding-top:10px;'>第 {exec_page} / {exec_total_pages} 頁，共 {exec_total_rows} 筆</div>",
        unsafe_allow_html=True,
    )

if execution_log_df.empty:
    st.warning("目前沒有執行紀錄")
else:
    st.dataframe(
        style_exec_log_display(exec_page_df),
        use_container_width=True,
        hide_index=True,
    )


# =========================
# 每日總覽留存紀錄（可刪除）
# =========================
st.header("當月每日業績總覽留存紀錄")

history_page_df, history_page, history_total_pages, history_total_rows = paginate_df(
    daily_history_df,
    page_key="daily_history_page",
    page_size=10,
)

hc1, hc2, hc3 = st.columns([1, 1, 3])
with hc1:
    if st.button("◀ 上一頁 ", key="history_prev"):
        st.session_state["daily_history_page"] = max(1, st.session_state["daily_history_page"] - 1)
        st.rerun()
with hc2:
    if st.button("下一頁 ▶ ", key="history_next"):
        st.session_state["daily_history_page"] = min(
            history_total_pages,
            st.session_state["daily_history_page"] + 1
        )
        st.rerun()
with hc3:
    st.markdown(
        f"<div style='font-size:18px;padding-top:10px;'>第 {history_page} / {history_total_pages} 頁，共 {history_total_rows} 筆</div>",
        unsafe_allow_html=True,
    )

if daily_history_df.empty:
    st.warning("目前沒有每日總覽留存紀錄")
else:
    st.markdown("### 勾選要刪除的紀錄")

    selectable_ids = history_page_df["id"].astype(str).tolist()
    default_ids = [x for x in st.session_state["delete_ids"] if x in selectable_ids]

    selected_ids = st.multiselect(
        "Choose options",
        options=selectable_ids,
        default=default_ids,
        key="history_multiselect",
        label_visibility="collapsed",
    )

    st.session_state["delete_ids"] = selected_ids

    if st.button("🗑 刪除勾選列", key="delete_history_rows"):
        deleted = delete_daily_history_rows(selected_ids)
        st.session_state["delete_ids"] = []
        reset_pages()
        st.success(f"已刪除 {deleted} 筆")
        st.rerun()

    display_history = history_page_df.copy()
    if "daily_json" in display_history.columns:
        display_history = display_history.drop(columns=["daily_json"])

    st.dataframe(
        style_daily_history_display(display_history),
        use_container_width=True,
        hide_index=True,
    )


# =========================
# 信件預覽
# =========================
st.header("信件內容預覽")

if not email_html:
    st.info("目前沒有信件預覽內容")
else:
    st.components.v1.html(email_html, height=520, scrolling=True)


# =========================
# Footer
# =========================
st.caption(f"Lemon Clean Scheduler Console · {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
