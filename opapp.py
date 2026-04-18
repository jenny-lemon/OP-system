import streamlit as st
from datetime import datetime

from dashboard_main import render_page

st.set_page_config(
    page_title="Jenny 排程控制台",
    page_icon="🍋",
    layout="wide",
)

TOP_PAGES = [
    ("主控表", "📋"),
    ("業績報表", "💹"),
    ("上下半月訂單", "🧾"),
    ("手動執行", "▶️"),
    ("Log 監控", "📄"),
    ("輸出檔案", "📂"),
    ("程式管理", "⚙️"),
    ("排程設定", "⏰"),
]

if "page" not in st.session_state:
    st.session_state.page = "主控表"

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
    background: #f3f4f6 !important;
    color: #0f172a;
}
[data-testid="stHeader"], [data-testid="stSidebar"] {
    display: none !important;
}
.block-container {
    padding-top: 0.8rem !important;
    max-width: 1900px !important;
}
.topbar {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 22px 28px;
    margin: 10px 0 22px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 4px 14px rgba(0,0,0,0.05);
}
.brand {
    display: flex;
    align-items: center;
    gap: 18px;
}
.brand-title {
    font-size: 24px;
    font-weight: 900;
    color: #0f172a;
}
.brand-divider {
    width: 1px;
    height: 34px;
    background: #cbd5e1;
}
.top-clock {
    font-family: monospace;
    font-size: 16px;
    font-weight: 700;
    color: #64748b;
}
.top-tabs {
    margin-bottom: 22px;
}
.top-tabs div[data-testid="stButton"] > button {
    height: 68px !important;
    font-size: 20px !important;
    font-weight: 800 !important;
    border-radius: 14px !important;
    border: none !important;
    background: #16243b !important;
    color: white !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
}
.top-tabs div[data-testid="stButton"] > button:hover {
    background: #1e3150 !important;
}
</style>
""", unsafe_allow_html=True)

now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
st.markdown(
    f"""
    <div class="topbar">
        <div class="brand">
            <div class="brand-title">🍋 Jenny 排程控制台</div>
            <div class="brand-divider"></div>
        </div>
        <div class="top-clock">🕒 {now_str}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

cols = st.columns(len(TOP_PAGES))
for i, (label, icon) in enumerate(TOP_PAGES):
    with cols[i]:
        st.markdown('<div class="top-tabs">', unsafe_allow_html=True)
        if st.button(f"{icon} {label}", key=f"top_{label}", use_container_width=True):
            st.session_state.page = label
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

render_page(st.session_state.page)
