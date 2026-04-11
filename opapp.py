"""
opapp.py  ──  營運報表控制台 v4
- 選單固定在頁面最上方（捲動後仍固定）
- 排程主控表每列有執行鍵
- 文字顏色修正，全部清晰可見
- 帳密從本機 accounts.py 讀取（加入 .gitignore 不上傳）
"""

from pathlib import Path
import subprocess
import sys
import json
import os
from datetime import datetime

import streamlit as st

# ── 帳密從 Streamlit secrets 讀取 ────────────────────────────────
def load_accounts():
    try:
        return {
            "台北": {
                "email": st.secrets["accounts"]["taipei"]["email"],
                "password": st.secrets["accounts"]["taipei"]["password"],
            },
            "台中": {
                "email": st.secrets["accounts"]["taichung"]["email"],
                "password": st.secrets["accounts"]["taichung"]["password"],
            },
        }
    except Exception:
        return {}

ACCOUNTS = load_accounts()

# ──────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_F   = BASE_DIR / "schedule_config.json"
RUNLOG_F   = BASE_DIR / "run_log.json"
OUTPUT_DIR.mkdir(exist_ok=True)

PYTHON_BIN = sys.executable or "python3"

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="營運報表控制台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────
# CSS  ── 固定頂部 + 清晰配色
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700;900&family=IBM+Plex+Mono:wght@400;600&display=swap');

/* ── Reset ── */
html, body, [class*="css"] {
    font-family: 'Noto Sans TC', sans-serif !important;
    background: #0d1117 !important;
    color: #e6edf3 !important;
}
#MainMenu, footer, header { visibility: hidden; }
print("DEBUG 排班統計表 main started")
section[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"]  { display: none !important; }

/* ══ 固定頂部 NavBar ══ */
.topnav {
    position: fixed;
    top: 0; left: 0; right: 0;
    z-index: 99999;
    background: #161b22;
    border-bottom: 2px solid #21262d;
    height: 54px;
    display: flex;
    align-items: center;
    padding: 0 20px;
    gap: 0;
    box-shadow: 0 4px 24px rgba(0,0,0,.6);
}
.nav-brand {
    font-size: 16px;
    font-weight: 900;
    color: #58a6ff;
    letter-spacing: .02em;
    margin-right: 24px;
    white-space: nowrap;
    flex-shrink: 0;
}
.nav-links {
    display: flex;
    gap: 2px;
    flex: 1;
}
.nav-link {
    font-size: 13px;
    font-weight: 700;
    color: #8b949e;
    padding: 7px 16px;
    border-radius: 6px;
    white-space: nowrap;
    cursor: pointer;
}
.nav-link.active {
    background: #1f6feb33;
    color: #79c0ff;
}
.nav-right {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 10px;
    flex-shrink: 0;
}
.region-chip {
    font-size: 12px;
    font-weight: 700;
    background: #1f6feb22;
    border: 1px solid #1f6feb66;
    color: #79c0ff;
    padding: 4px 14px;
    border-radius: 20px;
}
.region-chip.none {
    background: #21262d;
    border-color: #30363d;
    color: #6e7681;
}
.nav-time {
    font-size: 12px;
    color: #6e7681;
    font-family: 'IBM Plex Mono', monospace;
}

/* ══ 實際導覽按鈕（隱形，蓋在 nav 上） ══
   放在頁面正文頂端，用 margin-top 推到 nav 位置 */
.nav-real-btns {
    position: fixed;
    top: 0; left: 180px;
    z-index: 100000;
    display: flex;
    gap: 2px;
    height: 54px;
    align-items: center;
}
.nav-real-btns .stButton > button {
    background: transparent !important;
    border: none !important;
    color: transparent !important;
    width: 110px !important;
    height: 40px !important;
    cursor: pointer !important;
    padding: 0 !important;
}

/* ══ Body offset ══ */
.block-container {
    padding-top: 70px !important;
    padding-bottom: 3rem;
    max-width: 1400px;
}

/* ══ Cards ══ */
.card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 20px 24px 18px;
    margin-bottom: 18px;
}
.card-title {
    font-size: 12px;
    font-weight: 900;
    color: #79c0ff;
    text-transform: uppercase;
    letter-spacing: .12em;
    margin-bottom: 16px;
}

/* ══ Sched Table ══ */
.sched-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
}
.sched-table th {
    background: #1c2128;
    color: #8b949e;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .08em;
    padding: 9px 14px;
    text-align: left;
    border-bottom: 1px solid #30363d;
}
.sched-table td {
    padding: 10px 14px;
    border-bottom: 1px solid #21262d;
    color: #e6edf3;
    vertical-align: middle;
}
.sched-table tr:last-child td { border-bottom: none; }
.sched-table tr:hover td { background: #1c2128; }
.sched-table td.dim { color: #8b949e; }
.sched-table code {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11.5px !important;
    background: #1c2128 !important;
    color: #a5d6ff !important;
    padding: 2px 6px !important;
    border-radius: 4px !important;
}

/* ══ Badges ══ */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    font-family: 'IBM Plex Mono', monospace;
    white-space: nowrap;
}
.badge-ok   { background:#0d4429; color:#56d364; border:1px solid #2ea043; }
.badge-fail { background:#3d0000; color:#ff7b72; border:1px solid #b62324; }
.badge-wait { background:#1c2128; color:#8b949e; border:1px solid #30363d; }

/* ══ Buttons ══ */
.stButton > button {
    background: #21262d !important;
    color: #e6edf3 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    transition: all .15s ease !important;
}
.stButton > button:hover {
    background: #1f6feb33 !important;
    color: #79c0ff !important;
    border-color: #388bfd !important;
}
/* 執行按鈕特別樣式 */
button[kind="secondary"] {
    font-size: 12px !important;
    padding: 4px 10px !important;
}

/* ══ Metrics ══ */
[data-testid="metric-container"] {
    background: #161b22 !important;
    border: 1px solid #30363d !important;
    border-radius: 10px !important;
    padding: 16px 20px !important;
}
[data-testid="stMetricLabel"] { color: #8b949e !important; font-size: 12px !important; }
[data-testid="stMetricValue"] { color: #e6edf3 !important; font-size: 24px !important; font-weight: 900 !important; }

/* ══ Selectbox ══ */
.stSelectbox label { color: #c9d1d9 !important; font-size: 13px !important; font-weight: 600 !important; }
.stSelectbox > div > div {
    background: #1c2128 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
    color: #e6edf3 !important;
}

/* ══ Text inputs ══ */
.stTextInput label, .stTextArea label, .stCheckbox label {
    color: #c9d1d9 !important;
    font-size: 13px !important;
}
.stTextInput input, .stTextArea textarea {
    background: #1c2128 !important;
    color: #e6edf3 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
    font-size: 13px !important;
}

/* ══ Tabs ══ */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid #30363d !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #8b949e !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    padding: 8px 20px !important;
    border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
    color: #79c0ff !important;
    border-bottom: 2px solid #388bfd !important;
}

/* ══ Alert / Info ══ */
[data-testid="stAlert"] { border-radius: 8px !important; }
.stAlert p { color: #e6edf3 !important; }

/* ══ Expander ══ */
details > summary {
    color: #c9d1d9 !important;
    font-size: 13px !important;
    font-weight: 600 !important;
}
.streamlit-expanderHeader { color: #c9d1d9 !important; }

/* ══ Caption / small text ══ */
.stCaption, [data-testid="stCaptionContainer"] { color: #8b949e !important; }

/* ══ Divider ══ */
hr { border-color: #21262d !important; margin: 10px 0 18px 0 !important; }

/* ══ Code blocks ══ */
.stCode, [data-testid="stCodeBlock"] {
    background: #1c2128 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Session state
# ──────────────────────────────────────────────
for k, v in {"page": "排程主控表", "region": None, "result_label": None, "result_data": None}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ──────────────────────────────────────────────
# Config / RunLog
# ──────────────────────────────────────────────
DEFAULT_CONFIG = {
    "daily": [
        {"id": "d1", "label": "排班統計表",   "script": "排班統計表.py",   "args": [],       "schedule": "01:00", "all_regions": False},
        {"id": "d2", "label": "專員班表",     "script": "專員班表.py",     "args": [],       "schedule": "02:00", "all_regions": False},
        {"id": "d3", "label": "專員個資",     "script": "專員系統個資.py", "args": [],       "schedule": "02:30", "all_regions": False},
        {"id": "d4", "label": "當月次月訂單", "script": "當月次月訂單.py", "args": [],       "schedule": "08:00", "all_regions": True},
        {"id": "d5", "label": "業績報表 08",  "script": "業績報表.py",     "args": ["0800"], "schedule": "08:00", "all_regions": True},
        {"id": "d6", "label": "業績報表 18",  "script": "業績報表.py",     "args": ["1800"], "schedule": "18:00", "all_regions": True},
    ],
    "monthly": [
        {"id": "m1", "label": "上半月訂單",  "script": "上下半月訂單.py", "args": ["1"], "schedule": "月初01日", "all_regions": True},
        {"id": "m2", "label": "下半月訂單",  "script": "上下半月訂單.py", "args": ["2"], "schedule": "月中16日", "all_regions": True},
        {"id": "m3", "label": "已退款",      "script": "已退款.py",       "args": [],    "schedule": "月底",    "all_regions": True},
        {"id": "m4", "label": "預收",        "script": "預收.py",         "args": [],    "schedule": "月底",    "all_regions": False},
        {"id": "m5", "label": "儲值金結算",  "script": "儲值金結算.py",   "args": [],    "schedule": "月底",    "all_regions": False},
        {"id": "m6", "label": "儲值金預收",  "script": "儲值金預收.py",   "args": [],    "schedule": "月底",    "all_regions": False},
    ],
}

def load_config():
    if CONFIG_F.exists():
        try:
            return json.loads(CONFIG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    CONFIG_F.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

def load_runlog():
    if RUNLOG_F.exists():
        try:
            return json.loads(RUNLOG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_runlog(log):
    RUNLOG_F.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")

def rkey(job_id, region=None):
    return f"{region}__{job_id}" if region else job_id

def record_run(key, ok, stdout, stderr):
    log = load_runlog()
    log[key] = {
        "last_run": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ok": ok,
        "stdout": (stdout or "")[-1200:],
        "stderr": (stderr or "")[-1200:],
    }
    save_runlog(log)

# ──────────────────────────────────────────────
# Script runner
# ──────────────────────────────────────────────
import subprocess
import sys

def run_script(script, args=None, region=None):
    if args is None:
        args = []

    cmd = [sys.executable, script] + args

    print("🚀 RUN:", cmd)
    print("🌏 REGION:", region)

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300
        )

        return {
            "ok": result.returncode == 0,
            "cmd": " ".join(cmd),
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    except Exception as e:
        return {
            "ok": False,
            "cmd": " ".join(cmd),
            "stdout": "",
            "stderr": str(e),
        }
    
def do_run_job(job, region):
    """執行 job，支援全區依序。回傳 [(地區, result), ...]。"""
    REGION_ALL = "【全部地區依序】"
    targets = list(ACCOUNTS.keys()) if region == REGION_ALL else ([region] if region else [None])
    results = []
    for r in targets:
        res = run_script(job["script"], job.get("args", []), region=r)
        record_run(rkey(job["id"], r), res["ok"], res["stdout"], res["stderr"])
        results.append((r or "—", res))
    return results

def scan_output():
    files = []
    for p in sorted(OUTPUT_DIR.rglob("*")):
        if p.is_file():
            s = p.stat()
            files.append({
                "name":    p.name,
                "folder":  str(p.parent.relative_to(OUTPUT_DIR)) if p.parent != OUTPUT_DIR else "根目錄",
                "size_kb": round(s.st_size / 1024, 1),
                "mtime":   datetime.fromtimestamp(s.st_mtime).strftime("%Y-%m-%d %H:%M"),
            })
    return sorted(files, key=lambda x: x["mtime"], reverse=True)

def new_id(prefix="x"):
    return f"{prefix}{datetime.now().strftime('%f')}"

# ──────────────────────────────────────────────
# Load
# ──────────────────────────────────────────────
cfg    = load_config()
runlog = load_runlog()
REGIONS     = sorted(ACCOUNTS.keys()) if ACCOUNTS else []
REGION_ALL  = "【全部地區依序】"

st.caption(f"可用地區：{', '.join(REGIONS) if REGIONS else '尚未讀到帳密'}")

# ══════════════════════════════════════════════════════════════
# 固定頂部 NavBar（HTML 視覺）
# ══════════════════════════════════════════════════════════════
PAGES = ["排程主控表", "手動執行", "輸出報表", "腳本管理"]
cur   = st.session_state.page
rgn   = st.session_state.region

nav_items = "".join(
    f'<span class="nav-link{"  active" if p == cur else ""}">{p}</span>'
    for p in PAGES
)
chip_cls  = "region-chip" if rgn and rgn != REGION_ALL else "region-chip none"
chip_text = f"📍 {rgn}" if rgn else "📍 未選地區"

st.markdown(f"""
<div class="topnav">
  <div class="nav-brand">📊 營運報表</div>
  <div class="nav-links">{nav_items}</div>
  <div class="nav-right">
    <span class="{chip_cls}">{chip_text}</span>
    <span class="nav-time">{datetime.now().strftime('%m/%d %H:%M')}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# 實際可點擊的導覽按鈕（固定在頂端，透明蓋在 HTML nav 上）
with st.container():
    st.markdown('<div class="nav-real-btns">', unsafe_allow_html=True)
    nav_cols = st.columns(len(PAGES))
    for i, p in enumerate(PAGES):
        with nav_cols[i]:
            if st.button(p, key=f"nav_{p}"):
                st.session_state.page = p
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 地區選擇器（每頁頂端）
# ══════════════════════════════════════════════════════════════
def region_selector(allow_all=False):
    if not ACCOUNTS:
        st.error("⚠️ 找不到 Streamlit secrets 帳密，請到 App settings > Secrets 設定 accounts.taipei / accounts.taichung。")
        return None

    opts = ["（不指定地區）"]
    if allow_all:
        opts.append(REGION_ALL)
    opts += REGIONS

    cur_rgn = st.session_state.region
    idx     = opts.index(cur_rgn) if cur_rgn in opts else 0

    col_sel, col_info = st.columns([2, 5])
    with col_sel:
        sel = st.selectbox("📍 操作地區", opts, index=idx, key="rgn_widget")
    with col_info:
        if sel not in ("（不指定地區）", REGION_ALL):
            if sel in ACCOUNTS:
                st.info(f"✅ **{sel}** 帳號：`{ACCOUNTS[sel].get('email','—')}`　密碼已遮蔽")
            else:
                st.warning(f"⚠️ accounts.py 中找不到「{sel}」的帳密")
        elif sel == REGION_ALL:
            st.info(f"🌐 將依序執行所有地區：{', '.join(REGIONS)}")
        else:
            st.caption("未指定地區時，腳本不會收到帳密環境變數")

    new_val = None if sel == "（不指定地區）" else sel
    if new_val != st.session_state.region:
        st.session_state.region = new_val
        st.rerun()
    return new_val

# ══════════════════════════════════════════════════════════════
# 執行結果顯示區（所有頁面共用）
# ══════════════════════════════════════════════════════════════
result_box = st.empty()

def show_run_result(label, pairs):
    with result_box.container():
        st.markdown(f"**執行結果：{label}**")
        for rgn_name, res in pairs:
            prefix = f"`{rgn_name}` " if rgn_name != "—" else ""
            if res["ok"]:
                st.success(f"✅ {prefix}完成")
            else:
                st.error(f"❌ {prefix}失敗")

            st.write("cmd:")
            st.code(res.get("cmd", ""), language="bash")

            st.write("stdout:")
            st.code(res.get("stdout", "") or "(empty)")

            st.write("stderr:")
            st.code(res.get("stderr", "") or "(empty)")
# ══════════════════════════════════════════════════════════════
# PAGE 1 ── 排程主控表
# ══════════════════════════════════════════════════════════════
if st.session_state.page == "排程主控表":

    region = region_selector()
    st.markdown("")

    # 統計
    all_jobs   = cfg["daily"] + cfg["monthly"]
    ok_count   = sum(1 for j in all_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is True)
    fail_count = sum(1 for j in all_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False)
    wait_count = len(all_jobs) - ok_count - fail_count

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📋 排程總數", len(all_jobs))
    m2.metric("✅ 成功",    ok_count)
    m3.metric("❌ 失敗",    fail_count)
    m4.metric("⏳ 待執行",  wait_count)

    def render_section(jobs, section_title, group_key):
        st.markdown(f'<div class="card"><div class="card-title">{section_title}</div></div>', unsafe_allow_html=True)

        # 表頭
        hcols = st.columns([3, 2, 3, 2, 2, 2, 1])
        for txt, col in zip(["名稱", "排程時間", "腳本", "參數", "地區範圍", "狀態 / 最後執行", "執行"], hcols):
            col.markdown(f"<span style='font-size:11px;font-weight:700;color:#8b949e;text-transform:uppercase;letter-spacing:.06em'>{txt}</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin:4px 0 0 0'>", unsafe_allow_html=True)

        for idx, job in enumerate(jobs):
            k      = rkey(job["id"], region)
            entry  = runlog.get(k) or runlog.get(job["id"], {})
            ok     = entry.get("ok", None)
            last   = entry.get("last_run", "—")
            exists = (BASE_DIR / job["script"]).exists()
            multi  = "🌐 全區" if job.get("all_regions") else "單區"

            if ok is True:
                badge = '✅ 成功'
                badge_color = "#56d364"
            elif ok is False:
                badge = '❌ 失敗'
                badge_color = "#ff7b72"
            else:
                badge = '— 待執行'
                badge_color = "#8b949e"

            args_str = " ".join(job.get("args", [])) or "—"
            e_icon   = "🟢" if exists else "🔴"

            row = st.columns([3, 2, 3, 2, 2, 2, 1])
            row[0].markdown(f"**{job['label']}**")
            row[1].markdown(f"<span style='color:#c9d1d9;font-size:13px'>{job['schedule']}</span>", unsafe_allow_html=True)
            row[2].markdown(f"{e_icon} `{job['script']}`")
            row[3].markdown(f"<span style='color:#c9d1d9'>{args_str}</span>", unsafe_allow_html=True)
            row[4].markdown(f"<span style='color:#c9d1d9'>{multi}</span>", unsafe_allow_html=True)
            row[5].markdown(
                f"<span style='color:{badge_color};font-weight:700;font-size:13px'>{badge}</span>"
                f"<br><span style='color:#8b949e;font-size:11px;font-family:\"IBM Plex Mono\",monospace'>{last}</span>",
                unsafe_allow_html=True
            )
            with row[6]:
                if st.button("▶", key=f"dash_run_{group_key}_{job['id']}_{idx}", help=f"執行 {job['label']}"):
                    with st.spinner(f"執行 {job['label']} …"):
                        pairs = do_run_job(job, region)
                    show_run_result(job["label"], pairs)
                    st.rerun()

            st.markdown("<div style='border-bottom:1px solid #21262d;margin:2px 0'></div>", unsafe_allow_html=True)

    render_section(cfg["daily"],   "📅 每日排程", "daily")
    render_section(cfg["monthly"], "🗓️ 月排程",   "monthly")

    # 失敗詳情
    failures = [
        (j, runlog.get(rkey(j["id"], region)) or runlog.get(j["id"]))
        for j in (cfg["daily"] + cfg["monthly"])
        if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False
    ]
    if failures:
        st.markdown('<div class="card"><div class="card-title">⚠️ 失敗詳情</div>', unsafe_allow_html=True)
        for job, entry in failures:
            with st.expander(f"❌ {job['label']}  ·  {entry.get('last_run','—')}"):
                if entry.get("stderr"):
                    st.code(entry["stderr"], language="bash")
                if entry.get("stdout"):
                    st.code(entry["stdout"])
        st.markdown("</div>", unsafe_allow_html=True)

    col_r, _ = st.columns([1, 5])
    with col_r:
        if st.button("🔄 重新整理"):
            st.rerun()


# ══════════════════════════════════════════════════════════════
# PAGE 2 ── 手動執行
# ══════════════════════════════════════════════════════════════
elif st.session_state.page == "手動執行":

    region = region_selector(allow_all=True)

    def run_and_show(job):
        with st.spinner(f"執行 {job['label']} …"):
            pairs = do_run_job(job, region)
        show_run_result(job["label"], pairs)

    # ── 每日 ──
    st.markdown('<div class="card"><div class="card-title">📅 每日報表</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, job in enumerate(cfg["daily"]):
        with cols[i % 3]:
            if st.button(job["label"], key=f"man_{job['id']}"):
                run_and_show(job)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── 月報 ──
    st.markdown('<div class="card"><div class="card-title">🗓️ 月報表</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for i, job in enumerate(cfg["monthly"]):
        with cols[i % 3]:
            if st.button(job["label"], key=f"man_{job['id']}"):
                run_and_show(job)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── 批次 ──
    st.markdown('<div class="card"><div class="card-title">🚀 批次執行</div>', unsafe_allow_html=True)
    bc1, bc2 = st.columns(2)

    def batch(jobs, title):
        summary = []
        for job in jobs:
            with st.spinner(f"執行 {job['label']} …"):
                pairs = do_run_job(job, region)
            summary.append((job["label"], pairs))
        with result_box.container():
            st.markdown(f"**{title} 結果**")
            for label, pairs in summary:
                for rgn_name, res in pairs:
                    prefix = f"`{rgn_name}` " if rgn_name != "—" else ""
                    if res["ok"]:
                        st.success(f"✅ {prefix}{label}")
                    else:
                        st.error(f"❌ {prefix}{label}")
                        if res.get("stderr"):
                            st.code(res["stderr"])

    with bc1:
        if st.button("▶ 全部每日報表", key="batch_d"):
            batch(cfg["daily"], "每日批次")
    with bc2:
        if st.button("▶ 全部月報表", key="batch_m"):
            batch(cfg["monthly"], "月報批次")

    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE 3 ── 輸出報表
# ══════════════════════════════════════════════════════════════
elif st.session_state.page == "輸出報表":

    files   = scan_output()
    folders = sorted({f["folder"] for f in files})

    m1, m2, m3 = st.columns(3)
    m1.metric("📄 檔案總數", len(files))
    m2.metric("💾 總大小", f"{sum(f['size_kb'] for f in files):.1f} KB")
    m3.metric("📂 資料夾數", len(folders))

    st.markdown('<div class="card"><div class="card-title">📂 輸出目錄瀏覽</div>', unsafe_allow_html=True)
    st.caption(f"掃描路徑：`{OUTPUT_DIR}`")

    sel_folder = st.selectbox("篩選資料夾", ["（全部）"] + folders)
    filtered   = files if sel_folder == "（全部）" else [f for f in files if f["folder"] == sel_folder]

    if not filtered:
        st.info("output/ 目錄目前沒有檔案，執行報表後檔案將顯示於此。")
    else:
        hcols = st.columns([4, 2, 2, 1])
        for txt, col in zip(["檔名", "資料夾", "修改時間", "大小"], hcols):
            col.markdown(f"<span style='font-size:11px;font-weight:700;color:#8b949e;text-transform:uppercase;letter-spacing:.06em'>{txt}</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin:4px 0 8px 0'>", unsafe_allow_html=True)
        for f in filtered:
            row = st.columns([4, 2, 2, 1])
            row[0].markdown(f"📄 **{f['name']}**")
            row[1].markdown(f"<span style='color:#c9d1d9;font-size:13px'>{f['folder']}</span>", unsafe_allow_html=True)
            row[2].markdown(f"<span style='color:#8b949e;font-family:\"IBM Plex Mono\",monospace;font-size:12px'>{f['mtime']}</span>", unsafe_allow_html=True)
            row[3].markdown(f"<span style='color:#79c0ff;font-family:\"IBM Plex Mono\",monospace;font-size:12px'>{f['size_kb']} KB</span>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    col_r, _ = st.columns([1, 5])
    with col_r:
        if st.button("🔄 重新掃描"):
            st.rerun()


# ══════════════════════════════════════════════════════════════
# PAGE 4 ── 腳本管理
# ══════════════════════════════════════════════════════════════
elif st.session_state.page == "腳本管理":

    tab_d, tab_m, tab_edit = st.tabs(["📅 每日排程腳本", "🗓️ 月排程腳本", "📝 線上編輯腳本"])

    def job_editor(group_key, tab):
        with tab:
            jobs = cfg[group_key]

            st.markdown('<div class="card"><div class="card-title">現有腳本</div>', unsafe_allow_html=True)
            for idx, job in enumerate(jobs):
                exists = (BASE_DIR / job["script"]).exists()
                icon   = "🟢" if exists else "🔴"
                with st.expander(f"{icon} {job['label']}  ·  `{job['script']}`"):
                    c1, c2 = st.columns(2)
                    with c1:
                        new_label  = st.text_input("顯示名稱",           job["label"],                 key=f"lbl_{job['id']}")
                        new_script = st.text_input("腳本檔名",            job["script"],                key=f"scr_{job['id']}")
                    with c2:
                        new_args   = st.text_input("參數（空格分隔）",    " ".join(job.get("args",[])), key=f"arg_{job['id']}")
                        new_sched  = st.text_input("排程說明",            job.get("schedule",""),       key=f"sch_{job['id']}")
                    new_multi = st.checkbox("需跑全部地區", value=job.get("all_regions", False), key=f"mul_{job['id']}")

                    a1, a2, a3 = st.columns(3)
                    with a1:
                        if st.button("💾 儲存", key=f"save_{job['id']}"):
                            cfg[group_key][idx].update({
                                "label": new_label, "script": new_script,
                                "args": new_args.split() if new_args.strip() else [],
                                "schedule": new_sched, "all_regions": new_multi,
                            })
                            save_config(cfg)
                            st.success("已儲存")
                            st.rerun()
                    with a2:
                        if st.button("▶ 測試執行", key=f"test_{job['id']}"):
                            with st.spinner("執行中…"):
                                pairs = do_run_job(job, st.session_state.region)
                            show_run_result(job["label"], pairs)
                    with a3:
                        if st.button("🗑️ 刪除", key=f"del_{job['id']}"):
                            cfg[group_key].pop(idx)
                            save_config(cfg)
                            st.warning(f"已刪除：{job['label']}")
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            # ── 新增 ──
            st.markdown('<div class="card"><div class="card-title">➕ 新增腳本</div>', unsafe_allow_html=True)
            with st.form(key=f"add_{group_key}"):
                c1, c2 = st.columns(2)
                with c1:
                    f_label  = st.text_input("顯示名稱",        placeholder="例：新報表")
                    f_script = st.text_input("腳本檔名",        placeholder="例：新報表.py")
                with c2:
                    f_args   = st.text_input("參數（空格分隔）", placeholder="例：0800")
                    f_sched  = st.text_input("排程時間說明",     placeholder="例：09:00")
                f_multi = st.checkbox("需跑全部地區")
                if st.form_submit_button("➕ 新增"):
                    if not f_label or not f_script:
                        st.error("名稱與腳本為必填")
                    else:
                        cfg[group_key].append({
                            "id": new_id(group_key[0]),
                            "label": f_label, "script": f_script,
                            "args": f_args.split() if f_args.strip() else [],
                            "schedule": f_sched or "—",
                            "all_regions": f_multi,
                        })
                        save_config(cfg)
                        st.success(f"已新增：{f_label}")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    job_editor("daily",   tab_d)
    job_editor("monthly", tab_m)

    with tab_edit:
        st.markdown('<div class="card"><div class="card-title">📝 直接編輯腳本內容</div>', unsafe_allow_html=True)
        all_scripts = sorted({j["script"] for grp in ["daily","monthly"] for j in cfg[grp]})
        sel = st.selectbox("選擇腳本", ["（請選擇）"] + all_scripts)
        if sel and sel != "（請選擇）":
            spath  = BASE_DIR / sel
            code   = spath.read_text(encoding="utf-8") if spath.exists() else f"# {sel} 尚未建立\n"
            edited = st.text_area("腳本內容", code, height=420)
            if st.button("💾 寫入儲存"):
                spath.write_text(edited, encoding="utf-8")
                st.success(f"已寫入 {sel}")
        st.markdown("</div>", unsafe_allow_html=True)

        with st.expander("⚠️ 重置所有排程設定為預設值"):
            if st.button("🔄 確認重置"):
                save_config(DEFAULT_CONFIG)
                st.warning("已重置，請重新整理頁面")
                st.rerun()
