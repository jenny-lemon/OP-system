import streamlit as st
from datetime import datetime, timedelta, timezone
from dashboard_main import render_page

st.set_page_config(
    page_title="Jenny 排程控制台",
    page_icon="🍋",
    layout="wide",
)

TZ_TAIPEI = timezone(timedelta(hours=8))

TOP_PAGES = [
    ("主控表",    "📋"),
    ("業績報表",  "💹"),
    ("上下半月訂單", "🧾"),
    ("手動執行",  "▶️"),
    ("Log 監控", "📄"),
    ("輸出檔案",  "📂"),
    ("程式管理",  "⚙️"),
    ("排程設定",  "⏰"),
]

if "page" not in st.session_state:
    st.session_state.page = "主控表"

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

/* ─── Base ─── */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"] {
    background: #f0f2f6 !important;
    font-family: 'DM Sans', 'PingFang TC', 'Noto Sans TC', sans-serif !important;
    color: #1e293b !important;
}
[data-testid="stHeader"],
[data-testid="stSidebar"] { display: none !important; }
.block-container { padding: 0 2.4rem 4rem !important; max-width: 1480px !important; }

/* ─── Topbar ─── */
.topbar {
    background: #ffffff;
    margin: 0 -2.4rem;
    padding: 0 28px;
    border-bottom: 1.5px solid #e2e8f0;
    position: sticky; top: 0; z-index: 999;
    box-shadow: 0 2px 12px rgba(15,23,42,.07);
}
.topbar-inner { display: flex; align-items: center; height: 58px; }
.topbar-brand { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
.topbar-logo  { font-size: 22px; line-height: 1; }
.topbar-name  { font-size: 15px; font-weight: 700; color: #0f172a; letter-spacing: -.01em; white-space: nowrap; }
.topbar-div   { width: 1.5px; height: 24px; background: #e2e8f0; margin: 0 22px; flex-shrink: 0; }
.topbar-clock { font-size: 12px; color: #64748b; font-weight: 500; margin-left: auto;
                font-variant-numeric: tabular-nums; }

/* ─── Nav strip ─── */
.nav-strip {
    background: #ffffff;
    margin: 0 -2.4rem;
    padding: 0 16px;
    border-bottom: 1.5px solid #e2e8f0;
    box-shadow: 0 1px 4px rgba(15,23,42,.04);
}

/* Nav buttons — default */
.nav-wrap div[data-testid="stButton"] > button {
    height: 42px !important;
    padding: 0 16px !important;
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 2.5px solid transparent !important;
    background: transparent !important;
    color: #64748b !important;
    font-weight: 600 !important;
    font-size: 12.5px !important;
    box-shadow: none !important;
    white-space: nowrap !important;
    transition: color .15s !important;
}
.nav-wrap div[data-testid="stButton"] > button:hover {
    color: #1e293b !important;
    background: #f8fafc !important;
}
/* Nav buttons — active */
.nav-wrap.active div[data-testid="stButton"] > button {
    color: #2563eb !important;
    border-bottom: 2.5px solid #2563eb !important;
    background: transparent !important;
}

/* ─── Page header ─── */
.page-header {
    padding: 26px 0 18px;
    border-bottom: 1.5px solid #e2e8f0;
    margin-bottom: 26px;
    display: flex; align-items: flex-end; gap: 14px;
}
.page-title    { font-size: 23px; font-weight: 700; color: #0f172a; line-height: 1; letter-spacing: -.02em; }
.page-subtitle { font-size: 10.5px; font-weight: 700; letter-spacing: .15em; text-transform: uppercase;
                 color: #94a3b8; padding-bottom: 2px; }

/* ─── KPI cards ─── */
.kpi-row { display: flex; gap: 14px; margin-bottom: 26px; }
.kpi-card {
    flex: 1; background: #fff; border: 1.5px solid #e2e8f0; border-radius: 14px;
    padding: 18px 22px 16px; position: relative; overflow: hidden;
    box-shadow: 0 1px 4px rgba(15,23,42,.05), 0 4px 14px rgba(15,23,42,.04);
}
.kpi-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3.5px;
    border-radius: 14px 14px 0 0;
}
.kpi-card.blue::before  { background: linear-gradient(90deg,#2563eb,#60a5fa); }
.kpi-card.green::before { background: linear-gradient(90deg,#059669,#34d399); }
.kpi-card.amber::before { background: linear-gradient(90deg,#b45309,#fbbf24); }
.kpi-card.red::before   { background: linear-gradient(90deg,#dc2626,#f87171); }
.kpi-label { font-size: 10px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: #64748b; margin-bottom: 8px; }
.kpi-value { font-size: 36px; font-weight: 700; color: #0f172a; line-height: 1; letter-spacing: -.03em; font-variant-numeric: tabular-nums; }
.kpi-sub   { font-size: 12px; color: #64748b; font-weight: 500; margin-top: 6px; }

/* ─── Section card ─── */
.section-card {
    background: #fff; border: 1.5px solid #e2e8f0; border-radius: 14px;
    padding: 22px 24px 20px; margin-bottom: 18px;
    box-shadow: 0 1px 3px rgba(15,23,42,.04), 0 4px 14px rgba(15,23,42,.04);
}
.section-title {
    font-size: 12px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
    color: #2563eb; margin-bottom: 18px; padding-bottom: 12px;
    border-bottom: 1px solid #f1f5f9;
    display: flex; align-items: center; gap: 8px;
}

/* ─── Status badges ─── */
.badge {
    display: inline-flex; align-items: center; gap: 5px;
    font-size: 11.5px; font-weight: 600; padding: 3px 10px; border-radius: 20px; white-space: nowrap;
}
.badge::before { content: ''; width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
.b-green  { color: #065f46; background: #d1fae5; } .b-green::before  { background: #059669; }
.b-yellow { color: #78350f; background: #fef3c7; } .b-yellow::before { background: #d97706; }
.b-red    { color: #991b1b; background: #fee2e2; } .b-red::before    { background: #dc2626; }
.b-gray   { color: #475569; background: #f1f5f9; } .b-gray::before   { background: #94a3b8; }

/* ─── Run button (small, dark) ─── */
.run-btn div[data-testid="stButton"] > button {
    background: #1e293b !important; color: #f1f5f9 !important;
    border: none !important; border-radius: 7px !important;
    font-weight: 600 !important; font-size: 12px !important;
    padding: 4px 12px !important; height: 30px !important; min-height: 30px !important;
    box-shadow: none !important;
}
.run-btn div[data-testid="stButton"] > button:hover { background: #0f172a !important; }

/* ─── Save button (small, blue tint) ─── */
.save-btn div[data-testid="stButton"] > button {
    background: #eff6ff !important; color: #1d4ed8 !important;
    border: 1.5px solid #bfdbfe !important; border-radius: 7px !important;
    font-weight: 700 !important; font-size: 12px !important;
    padding: 3px 10px !important; height: 30px !important; min-height: 30px !important;
    box-shadow: none !important;
}
.save-btn div[data-testid="stButton"] > button:hover { background: #dbeafe !important; }

/* ─── Exec result panel ─── */
.exec-panel {
    background: #fff; border: 1.5px solid #e2e8f0; border-radius: 12px;
    padding: 16px 20px; margin-top: 14px;
    box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
.exec-panel-title {
    font-size: 13px; font-weight: 700; color: #0f172a;
    margin-bottom: 10px; display: flex; align-items: center; gap: 8px;
}
.exec-label {
    font-size: 9.5px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
    color: #94a3b8; margin: 10px 0 5px;
}

/* ─── Log box ─── */
.log-box {
    background: #0d1117; border: 1px solid #1e2d3d; border-radius: 10px;
    padding: 14px 18px;
    font-family: 'DM Mono', 'Menlo', monospace; font-size: 12.5px;
    line-height: 1.75; white-space: pre-wrap; word-break: break-all;
    max-height: 420px; overflow: auto;
}
.log-err    { color: #f87171; display: block; }
.log-ok     { color: #4ade80; display: block; }
.log-warn   { color: #fbbf24; display: block; }
.log-info   { color: #60a5fa; display: block; }
.log-normal { color: #94a3b8; display: block; }
.log-meta   { font-size: 11.5px; color: #64748b; font-weight: 500; margin-bottom: 8px; }

/* ─── Next-run chip ─── */
.next-run {
    display: inline-flex; align-items: center; gap: 5px;
    background: #f8fafc; border: 1.5px solid #e2e8f0; border-radius: 8px;
    padding: 4px 10px; font-size: 11.5px; font-weight: 600; color: #475569;
}

/* ─── Empty state ─── */
.empty-state {
    text-align: center; padding: 32px 20px; color: #94a3b8; font-size: 13px; font-weight: 500;
    background: #f8fafc; border-radius: 10px; border: 1.5px dashed #e2e8f0;
}
.empty-state .icon { font-size: 28px; display: block; margin-bottom: 8px; }

/* ─── Divider ─── */
.divider { height: 1.5px; background: #e2e8f0; margin: 20px 0; }

/* ─── Streamlit component overrides ─── */
div[data-testid="stButton"] > button {
    background: #1e293b !important; color: #f8fafc !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important; font-size: 13px !important;
    padding: 8px 18px !important; box-shadow: 0 1px 3px rgba(15,23,42,.12) !important;
}
div[data-testid="stButton"] > button:hover { background: #0f172a !important; }

div[data-testid="stSelectbox"] > div > div,
div[data-testid="stTextInput"] > div > div > input {
    background: #fff !important; border: 1.5px solid #cbd5e1 !important;
    border-radius: 8px !important; color: #1e293b !important; font-size: 13.5px !important;
}
div[data-testid="stSelectbox"] label,
div[data-testid="stTextInput"] label,
div[data-testid="stTextArea"] label,
div[data-testid="stRadio"] label {
    color: #374151 !important; font-size: 13px !important; font-weight: 600 !important;
}
div[data-testid="stTextArea"] textarea {
    border: 1.5px solid #cbd5e1 !important; border-radius: 8px !important;
    font-family: 'DM Mono', monospace !important; font-size: 12.5px !important;
    color: #1e293b !important; background: #fafafa !important;
}
div[data-testid="stMetric"] {
    background: #fff; border-radius: 12px; padding: 16px 18px;
    border: 1.5px solid #e2e8f0; box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
div[data-testid="stMetric"] label {
    color: #475569 !important; font-size: 12px !important; font-weight: 600 !important;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #0f172a !important; font-size: 28px !important; font-weight: 700 !important;
}
div[data-testid="stDataFrame"] {
    border-radius: 10px !important; overflow: hidden !important; border: 1.5px solid #e2e8f0 !important;
}
div[data-testid="stAlert"] {
    border-radius: 10px !important; font-size: 13px !important; font-weight: 500 !important;
}
.stCaption, div[data-testid="stCaption"] {
    color: #64748b !important; font-size: 12px !important; font-weight: 500 !important;
}
div[data-testid="stCheckbox"] label {
    color: #374151 !important; font-size: 13px !important; font-weight: 500 !important;
}
h3 { color: #0f172a !important; font-size: 16px !important; font-weight: 700 !important; }

/* ─── Footer ─── */
.footer-cap {
    text-align: center; font-size: 11px; color: #94a3b8; font-weight: 500;
    padding-top: 28px; border-top: 1.5px solid #e2e8f0; margin-top: 32px;
}
</style>
""", unsafe_allow_html=True)

# ── Topbar ────────────────────────────────────────────────────────────────────
now_str = datetime.now(TZ_TAIPEI).strftime("%Y/%m/%d  %H:%M")
st.markdown(
    f"""<div class="topbar">
      <div class="topbar-inner">
        <div class="topbar-brand">
          <span class="topbar-logo">🍋</span>
          <span class="topbar-name">Jenny 排程控制台</span>
        </div>
        <div class="topbar-div"></div>
        <div class="topbar-clock">🕐 {now_str}</div>
      </div>
    </div>""",
    unsafe_allow_html=True,
)

# ── Navigation strip ──────────────────────────────────────────────────────────
st.markdown('<div class="nav-strip">', unsafe_allow_html=True)
nav_cols = st.columns(len(TOP_PAGES))
for i, (label, icon) in enumerate(TOP_PAGES):
    active = st.session_state.page == label
    with nav_cols[i]:
        cls = "nav-wrap active" if active else "nav-wrap"
        st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)
        if st.button(f"{icon} {label}", key=f"nav_{label}"):
            st.session_state.page = label
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# ── Render page content ───────────────────────────────────────────────────────
render_page(st.session_state.page)
