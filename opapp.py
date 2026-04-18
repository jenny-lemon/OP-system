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
    ("主控表",       "📋"),
    ("業績報表",     "💹"),
    ("上下半月訂單", "🧾"),
    ("手動執行",     "▶️"),
    ("Log 監控",    "📄"),
    ("輸出檔案",     "📂"),
    ("程式管理",     "⚙️"),
    ("排程設定",     "⏰"),
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
    font-family: 'DM Sans','PingFang TC','Noto Sans TC',sans-serif !important;
    color: #1e293b !important;
}
[data-testid="stHeader"],
[data-testid="stSidebar"] { display: none !important; }
.block-container { padding: 0 2.4rem 4rem !important; max-width: 1480px !important; }

/* ─── Topbar ─── */
.topbar {
    background: #fff;
    margin: 0 -2.4rem;
    padding: 0 28px;
    border-bottom: 1px solid #e8ecf0;
    position: sticky; top: 0; z-index: 999;
    box-shadow: 0 1px 8px rgba(15,23,42,.06);
}
.topbar-inner { display: flex; align-items: center; height: 52px; }
.topbar-brand { display: flex; align-items: center; gap: 9px; flex-shrink: 0; }
.topbar-logo  { font-size: 20px; line-height: 1; }
.topbar-name  { font-size: 14px; font-weight: 700; color: #0f172a; letter-spacing: -.01em; white-space: nowrap; }
.topbar-sep   { width: 1px; height: 20px; background: #dde2e8; margin: 0 18px; flex-shrink: 0; }
.topbar-clock { font-size: 12px; color: #64748b; font-weight: 500; margin-left: auto; font-variant-numeric: tabular-nums; }

/* ─── Nav strip ─── */
.nav-strip {
    background: #fff;
    margin: 0 -2.4rem;
    padding: 0 12px;
    border-bottom: 1px solid #e8ecf0;
}

/* ─── Nav buttons — ALWAYS override global button style ─── */
html body .nav-wrap div[data-testid="stButton"] > button,
html body .nav-wrap div[data-testid="stButton"] > button:focus,
html body .nav-wrap div[data-testid="stButton"] > button:active {
    height: 40px !important;
    padding: 0 14px !important;
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
    color: #64748b !important;
    font-weight: 600 !important;
    font-size: 12px !important;
    box-shadow: none !important;
    white-space: nowrap !important;
    letter-spacing: 0 !important;
}
html body .nav-wrap div[data-testid="stButton"] > button:hover {
    color: #1e293b !important;
    background: #f4f6f9 !important;
    border-bottom: 2px solid transparent !important;
}
html body .nav-wrap.active div[data-testid="stButton"] > button {
    color: #2563eb !important;
    background: #eff6ff !important;
    border-bottom: 2px solid #2563eb !important;
}

/* ─── Page header ─── */
.page-header {
    padding: 22px 0 16px;
    border-bottom: 1px solid #e8ecf0;
    margin-bottom: 22px;
    display: flex; align-items: flex-end; gap: 12px;
}
.page-title    { font-size: 22px; font-weight: 700; color: #0f172a; line-height: 1; letter-spacing: -.02em; }
.page-subtitle { font-size: 10px; font-weight: 700; letter-spacing: .14em; text-transform: uppercase; color: #94a3b8; padding-bottom: 1px; }

/* ─── KPI cards ─── */
.kpi-row { display: flex; gap: 12px; margin-bottom: 22px; }
.kpi-card {
    flex: 1; background: #fff; border: 1px solid #e8ecf0; border-radius: 12px;
    padding: 16px 20px 14px; position: relative; overflow: hidden;
    box-shadow: 0 1px 3px rgba(15,23,42,.04), 0 3px 10px rgba(15,23,42,.04);
}
.kpi-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; border-radius: 12px 12px 0 0;
}
.kpi-card.blue::before  { background: linear-gradient(90deg,#2563eb,#60a5fa); }
.kpi-card.green::before { background: linear-gradient(90deg,#059669,#34d399); }
.kpi-card.amber::before { background: linear-gradient(90deg,#b45309,#fbbf24); }
.kpi-card.red::before   { background: linear-gradient(90deg,#dc2626,#f87171); }
.kpi-label { font-size: 9.5px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: #64748b; margin-bottom: 7px; }
.kpi-value { font-size: 34px; font-weight: 700; color: #0f172a; line-height: 1; letter-spacing: -.03em; font-variant-numeric: tabular-nums; }
.kpi-sub   { font-size: 11.5px; color: #64748b; font-weight: 500; margin-top: 5px; }

/* ─── Section card ─── */
.section-card {
    background: #fff; border: 1px solid #e8ecf0; border-radius: 12px;
    padding: 20px 22px 18px; margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(15,23,42,.04), 0 3px 10px rgba(15,23,42,.04);
}
.section-title {
    font-size: 11px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
    color: #2563eb; margin-bottom: 16px; padding-bottom: 12px;
    border-bottom: 1px solid #f1f5f9;
    display: flex; align-items: center; gap: 7px;
}

/* ─── Status badges ─── */
.badge {
    display: inline-flex; align-items: center; gap: 5px;
    font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; white-space: nowrap;
}
.badge::before { content:''; width:5px; height:5px; border-radius:50%; flex-shrink:0; }
.b-green  { color:#065f46; background:#d1fae5; } .b-green::before  { background:#059669; }
.b-yellow { color:#78350f; background:#fef3c7; } .b-yellow::before { background:#d97706; }
.b-red    { color:#991b1b; background:#fee2e2; } .b-red::before    { background:#dc2626; }
.b-gray   { color:#475569; background:#f1f5f9; } .b-gray::before   { background:#94a3b8; }
.b-blue   { color:#1d4ed8; background:#dbeafe; } .b-blue::before   { background:#2563eb; }

/* ─── Run button ─── */
html body .run-btn div[data-testid="stButton"] > button {
    background: #1e293b !important; color: #f1f5f9 !important;
    border: none !important; border-radius: 6px !important;
    font-weight: 600 !important; font-size: 12px !important;
    padding: 3px 11px !important; height: 28px !important; min-height: 28px !important;
    box-shadow: none !important;
}
html body .run-btn div[data-testid="stButton"] > button:hover { background: #0f172a !important; }

/* ─── Save button ─── */
html body .save-btn div[data-testid="stButton"] > button {
    background: #f0f9ff !important; color: #0369a1 !important;
    border: 1px solid #bae6fd !important; border-radius: 6px !important;
    font-weight: 700 !important; font-size: 12px !important;
    padding: 2px 9px !important; height: 28px !important; min-height: 28px !important;
    box-shadow: none !important;
}
html body .save-btn div[data-testid="stButton"] > button:hover { background: #e0f2fe !important; }

/* ─── Inline task result (success/fail tag under each row) ─── */
.task-result-row {
    padding: 6px 4px 10px;
    border-bottom: 1px solid #f1f5f9;
    margin-bottom: 4px;
}
.task-result-ok   { background: #f0fdf4; border-radius: 8px; padding: 8px 14px; font-size: 12.5px; color: #166534; font-weight: 500; }
.task-result-fail { background: #fef2f2; border-radius: 8px; padding: 8px 14px; font-size: 12.5px; color: #991b1b; font-weight: 500; }

/* ─── Exec log panel ─── */
.exec-panel {
    background: #fff; border: 1px solid #e8ecf0; border-left: 3px solid #2563eb;
    border-radius: 10px; padding: 14px 18px; margin-top: 10px;
    box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
.exec-panel.ok   { border-left-color: #059669; }
.exec-panel.fail { border-left-color: #dc2626; }
.exec-panel-title { font-size: 13px; font-weight: 700; color: #0f172a; margin-bottom: 8px; display:flex; align-items:center; gap:8px; }
.exec-label { font-size: 9.5px; font-weight: 700; letter-spacing:.1em; text-transform:uppercase; color:#94a3b8; margin: 10px 0 5px; }

/* ─── Log box ─── */
.log-box {
    background: #0d1117; border: 1px solid #1e2d3d; border-radius: 9px;
    padding: 12px 16px;
    font-family: 'DM Mono','Menlo',monospace; font-size: 12px;
    line-height: 1.75; white-space: pre-wrap; word-break: break-all;
    max-height: 380px; overflow: auto;
}
.log-err    { color: #f87171; display: block; }
.log-ok     { color: #4ade80; display: block; }
.log-warn   { color: #fbbf24; display: block; }
.log-info   { color: #60a5fa; display: block; }
.log-normal { color: #94a3b8; display: block; }
.log-meta   { font-size: 11px; color: #64748b; font-weight: 500; margin-bottom: 7px; }

/* ─── Next-run chip ─── */
.next-run {
    display: inline-flex; align-items: center; gap: 5px;
    background: #f8fafc; border: 1px solid #e8ecf0; border-radius: 7px;
    padding: 3px 9px; font-size: 11px; font-weight: 600; color: #475569;
}

/* ─── Command preview ─── */
.cmd-preview {
    background: #1e293b; color: #94a3b8; border-radius: 9px;
    padding: 10px 16px; font-family: 'DM Mono','Menlo',monospace; font-size: 12px;
    margin: 10px 0 14px; word-break: break-all;
}
.cmd-preview .cmd-hl { color: #60a5fa; }
.cmd-preview .cmd-arg { color: #a3e635; }
.cmd-preview .cmd-city { color: #fb923c; }

/* ─── Empty state ─── */
.empty-state {
    text-align: center; padding: 28px 20px; color: #94a3b8; font-size: 12.5px; font-weight: 500;
    background: #f8fafc; border-radius: 9px; border: 1px dashed #dde2e8;
}
.empty-state .icon { font-size: 26px; display: block; margin-bottom: 7px; }

/* ─── Date range chip ─── */
.date-chip {
    display: inline-flex; align-items: center; gap: 6px;
    background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px;
    padding: 5px 12px; font-size: 12.5px; font-weight: 600; color: #0369a1; margin: 8px 0 12px;
}

/* ─── Streamlit overrides ─── */
div[data-testid="stButton"] > button {
    background: #1e293b !important; color: #f8fafc !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important; font-size: 13px !important;
    padding: 8px 18px !important; box-shadow: 0 1px 3px rgba(15,23,42,.12) !important;
}
div[data-testid="stButton"] > button:hover { background: #0f172a !important; }

div[data-testid="stSelectbox"] > div > div,
div[data-testid="stTextInput"] > div > div > input {
    background: #fff !important; border: 1px solid #d1d9e0 !important;
    border-radius: 8px !important; color: #1e293b !important; font-size: 13.5px !important;
}
div[data-testid="stSelectbox"] label,
div[data-testid="stTextInput"] label,
div[data-testid="stTextArea"] label,
div[data-testid="stRadio"] label {
    color: #374151 !important; font-size: 13px !important; font-weight: 600 !important;
}
div[data-testid="stTextArea"] textarea {
    border: 1px solid #d1d9e0 !important; border-radius: 8px !important;
    font-family: 'DM Mono',monospace !important; font-size: 12.5px !important;
    color: #1e293b !important; background: #fafafa !important;
}
div[data-testid="stMetric"] {
    background: #fff; border-radius: 11px; padding: 14px 16px;
    border: 1px solid #e8ecf0; box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
div[data-testid="stMetric"] label { color: #475569 !important; font-size: 12px !important; font-weight: 600 !important; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #0f172a !important; font-size: 26px !important; font-weight: 700 !important; }
div[data-testid="stDataFrame"] { border-radius: 9px !important; overflow: hidden !important; border: 1px solid #e8ecf0 !important; }
div[data-testid="stAlert"] { border-radius: 9px !important; font-size: 13px !important; font-weight: 500 !important; }
.stCaption, div[data-testid="stCaption"] { color: #64748b !important; font-size: 11.5px !important; font-weight: 500 !important; }
div[data-testid="stCheckbox"] label { color: #374151 !important; font-size: 13px !important; font-weight: 500 !important; }
h3 { color: #0f172a !important; font-size: 15px !important; font-weight: 700 !important; }

/* ─── Footer ─── */
.footer-cap {
    text-align: center; font-size: 11px; color: #94a3b8; font-weight: 500;
    padding-top: 24px; border-top: 1px solid #e8ecf0; margin-top: 28px;
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
        <div class="topbar-sep"></div>
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

# ── Render page ───────────────────────────────────────────────────────────────
render_page(st.session_state.page)
