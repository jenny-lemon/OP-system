"""
opapp.py  ──  營運報表控制台 v5.1
改版：
  - 配色改為舒適中灰底，文字對比全面提升
  - 時區改為台灣台北時區 (Asia/Taipei)
  - 每日排程新增「業績報表」
  - NavBar 重構，不再疊透明層
"""

import json
import os
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st

TZ = ZoneInfo("Asia/Taipei")

def now_tw():
    return datetime.now(TZ)

# ══════════════════════════════════════════════════
# 帳密從 Streamlit secrets 讀取
# ══════════════════════════════════════════════════
def load_accounts():
    try:
        return {
            "台北": {
                "email":    st.secrets["accounts"]["taipei"]["email"],
                "password": st.secrets["accounts"]["taipei"]["password"],
            },
            "台中": {
                "email":    st.secrets["accounts"]["taichung"]["email"],
                "password": st.secrets["accounts"]["taichung"]["password"],
            },
        }
    except Exception:
        return {}


ACCOUNTS = load_accounts()

# ══════════════════════════════════════════════════
# Paths
# ══════════════════════════════════════════════════
BASE_DIR   = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR    = BASE_DIR / "logs"
CONFIG_F   = BASE_DIR / "schedule_config.json"
RUNLOG_F   = BASE_DIR / "run_log.json"
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# ══════════════════════════════════════════════════
# Page config
# ══════════════════════════════════════════════════
st.set_page_config(
    page_title="營運報表控制台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════
# CSS  ── 中灰底、高對比文字
# ══════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=DM+Mono:wght@400;500&family=Noto+Sans+TC:wght@400;500;700&display=swap');

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'Noto Sans TC', sans-serif !important;
    background: #1c2333 !important;
    color: #d4dce8 !important;
}
#MainMenu, footer, header { visibility: hidden; }
section[data-testid="stSidebar"],
[data-testid="collapsedControl"] { display: none !important; }

.block-container {
    padding-top: 0 !important;
    padding-bottom: 3rem;
    max-width: 1360px;
}

/* ── Top strip ── */
.topstrip {
    background: #141b27;
    border-bottom: 1px solid #2a3650;
    padding: 0 28px;
    margin: 0 -1rem 0;
    display: flex;
    align-items: center;
    height: 54px;
    position: sticky;
    top: 0;
    z-index: 9999;
    box-shadow: 0 2px 16px rgba(0,0,0,.35);
}
.topstrip-brand {
    font-family: 'Syne', sans-serif;
    font-size: 15px;
    font-weight: 800;
    color: #7db8f7;
    margin-right: 20px;
    white-space: nowrap;
    flex-shrink: 0;
}
.topstrip-time {
    margin-left: auto;
    font-family: 'DM Mono', monospace;
    font-size: 12.5px;
    color: #6b7fa0;
    flex-shrink: 0;
}
.rgn-chip {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    background: rgba(125,184,247,.1);
    border: 1px solid rgba(125,184,247,.25);
    color: #7db8f7;
    padding: 3px 12px;
    border-radius: 20px;
    margin-left: 8px;
}
.rgn-chip.none {
    background: rgba(100,116,139,.1);
    border-color: rgba(100,116,139,.25);
    color: #6b7fa0;
}

/* ── Nav bar ── */
div[data-testid="stHorizontalBlock"]:has(div.nav-wrap) {
    background: #141b27 !important;
    border-bottom: 1px solid #2a3650 !important;
    padding: 0 16px !important;
    margin: 0 -1rem 28px !important;
    gap: 0 !important;
}
.nav-wrap div[data-testid="stButton"] > button {
    height: 44px !important;
    padding: 0 20px !important;
    background: transparent !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    border-radius: 0 !important;
    color: #5a7090 !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 13px !important;
    box-shadow: none !important;
    transition: color .15s, border-color .15s !important;
}
.nav-wrap div[data-testid="stButton"] > button:hover {
    color: #a8c8f0 !important;
    border-bottom-color: #4a90d9 !important;
}
.nav-wrap.active div[data-testid="stButton"] > button {
    color: #7db8f7 !important;
    border-bottom-color: #4a8fd4 !important;
}

/* ── KPI Cards ── */
.kpi-row { display:flex; gap:14px; margin-bottom:24px; }
.kpi {
    flex:1; background:#222d42; border:1px solid #2e3f5c;
    border-radius:14px; padding:20px 24px; position:relative; overflow:hidden;
}
.kpi::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; }
.kpi.blue::before  { background: linear-gradient(90deg,#2563eb,#7db8f7); }
.kpi.green::before { background: linear-gradient(90deg,#059669,#34d399); }
.kpi.red::before   { background: linear-gradient(90deg,#b91c1c,#f87171); }
.kpi.amber::before { background: linear-gradient(90deg,#b45309,#fbbf24); }
.kpi-label { font-family:'DM Mono',monospace; font-size:10.5px; font-weight:500; letter-spacing:.12em; text-transform:uppercase; color:#5a7090; margin-bottom:8px; }
.kpi-value { font-family:'Syne',sans-serif; font-size:34px; font-weight:800; color:#e2eaf6; line-height:1; }
.kpi-sub   { font-size:12px; color:#6b7fa0; margin-top:5px; }

/* ── Panel ── */
.panel {
    background: #222d42;
    border: 1px solid #2e3f5c;
    border-radius: 14px;
    padding: 22px 24px 20px;
    margin-bottom: 20px;
}
.panel-head {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 18px;
    padding-bottom: 14px;
    border-bottom: 1px solid #2a3650;
}
.panel-tag {
    font-family: 'DM Mono', monospace;
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: #7db8f7;
    background: rgba(125,184,247,.1);
    border: 1px solid rgba(125,184,247,.22);
    border-radius: 6px;
    padding: 4px 12px;
}
.panel-note { font-size: 12px; color: #5a7090; margin-left: auto; }

/* ── Page titles ── */
.pg-title {
    font-family: 'Syne', sans-serif;
    font-size: 24px;
    font-weight: 800;
    color: #e2eaf6;
    margin-bottom: 2px;
    letter-spacing: -.01em;
}
.pg-sub {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: #3a4f6a;
    letter-spacing: .14em;
    margin-bottom: 24px;
}

/* ── Badges ── */
.badge {
    display: inline-flex; align-items: center; gap: 5px;
    font-family: 'DM Mono', monospace; font-size: 11.5px; font-weight: 500;
    padding: 3px 11px; border-radius: 20px; white-space: nowrap;
}
.badge .dot { width:6px; height:6px; border-radius:50%; flex-shrink:0; }
.badge.green { background:rgba(5,150,105,.15); border:1px solid rgba(52,211,153,.3); color:#4ade80; }
.badge.green .dot  { background:#34d399; }
.badge.red   { background:rgba(185,28,28,.15); border:1px solid rgba(248,113,113,.3); color:#f87171; }
.badge.red   .dot  { background:#ef4444; }
.badge.amber { background:rgba(180,83,9,.15); border:1px solid rgba(251,191,36,.3); color:#fbbf24; }
.badge.amber .dot  { background:#f59e0b; }
.badge.gray  { background:rgba(51,65,85,.4); border:1px solid rgba(100,116,139,.35); color:#8fa3be; }
.badge.gray  .dot  { background:#64748b; }

/* ── Log box ── */
.logbox {
    background: #161e2e;
    border: 1px solid #2a3650;
    border-radius: 12px;
    padding: 16px 20px;
    font-family: 'DM Mono', monospace;
    font-size: 12.5px;
    line-height: 1.75;
    white-space: pre-wrap;
    word-break: break-all;
    max-height: 520px;
    overflow-y: auto;
}
.le { color: #f87171; display:block; }
.lo { color: #4ade80; display:block; }
.lw { color: #fbbf24; display:block; }
.li { color: #60a5fa; display:block; }
.ln { color: #8fa3be; display:block; }

/* ── Next run ── */
.next-run {
    background: rgba(37,99,235,.1);
    border: 1px solid rgba(37,99,235,.25);
    border-radius: 10px;
    padding: 10px 16px;
    font-family: 'DM Mono', monospace;
    font-size: 12.5px;
    color: #7db8f7;
    margin-top: 10px;
    display: flex; align-items: center; gap: 8px;
}

/* ── Streamlit overrides ── */
.stButton > button {
    background: #2a3650 !important;
    color: #c8d8ec !important;
    border: 1px solid #3a4f6a !important;
    border-radius: 8px !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    transition: all .15s !important;
}
.stButton > button:hover {
    background: rgba(37,99,235,.2) !important;
    color: #7db8f7 !important;
    border-color: #4a8fd4 !important;
}
.stSelectbox label, .stTextInput label, .stTextArea label,
.stCheckbox label, .stMultiSelect label {
    color: #8fa3be !important;
    font-size: 13px !important;
    font-weight: 600 !important;
}
.stSelectbox > div > div,
.stTextInput input,
.stTextArea textarea {
    background: #161e2e !important;
    color: #d4dce8 !important;
    border: 1px solid #2e3f5c !important;
    border-radius: 8px !important;
    font-size: 13px !important;
}
[data-testid="stMetricLabel"]  { color: #6b7fa0 !important; font-size: 12px !important; }
[data-testid="stMetricValue"]  { color: #e2eaf6 !important; font-size: 26px !important; font-weight: 800 !important; }
[data-testid="metric-container"] {
    background: #222d42 !important;
    border: 1px solid #2e3f5c !important;
    border-radius: 12px !important;
    padding: 16px 20px !important;
}
.stTabs [data-baseweb="tab-list"] { background:transparent !important; border-bottom:1px solid #2a3650 !important; gap:0 !important; }
.stTabs [data-baseweb="tab"] { background:transparent !important; color:#5a7090 !important; font-size:13px !important; font-weight:700 !important; padding:8px 22px !important; border-radius:0 !important; }
.stTabs [aria-selected="true"] { color:#7db8f7 !important; border-bottom:2px solid #4a8fd4 !important; }
.stCaption, [data-testid="stCaptionContainer"] { color: #6b7fa0 !important; font-size: 12px !important; }
details > summary, .streamlit-expanderHeader { color: #a0b4cc !important; font-size: 13px !important; font-weight: 600 !important; }
hr { border-color: #2a3650 !important; margin: 8px 0 16px !important; }
[data-testid="stCodeBlock"] { background: #161e2e !important; border: 1px solid #2a3650 !important; border-radius: 8px !important; }
[data-testid="stAlert"] { border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# Session state
# ══════════════════════════════════════════════════
for k, v in {"page": "排程主控表", "region": None}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════
# Config / RunLog
# ══════════════════════════════════════════════════
DEFAULT_CONFIG = {
    "daily": [
        {"id": "d1", "label": "排班統計表",  "script": "schedule_report.py", "args": [], "schedule": "01:00", "all_regions": False},
        {"id": "d2", "label": "專員班表",    "script": "staff_schedule.py",  "args": [], "schedule": "02:00", "all_regions": False},
        {"id": "d3", "label": "專員個資",    "script": "staff_info.py",      "args": [], "schedule": "02:30", "all_regions": False},
        {"id": "d4", "label": "當月次月訂單","script": "orders_report.py",   "args": [], "schedule": "08:00", "all_regions": True},
        {"id": "d5", "label": "業績報表",    "script": "performance_report.py", "args": [], "schedule": "08:00", "all_regions": True},
    ],
    "monthly": [
        {"id": "m1", "label": "上半月訂單",  "script": "上下半月訂單.py", "args": ["1"], "schedule": "月初01日", "all_regions": True},
        {"id": "m2", "label": "下半月訂單",  "script": "上下半月訂單.py", "args": ["2"], "schedule": "月中16日", "all_regions": True},
        {"id": "m3", "label": "已退款",      "script": "已退款.py",       "args": [],    "schedule": "月底",    "all_regions": True},
        {"id": "m4", "label": "預收",        "script": "預收.py",         "args": [],    "schedule": "月底",    "all_regions": False},
        {"id": "m5", "label": "儲值金結算",  "script": "儲值金結算.py",   "args": [],    "schedule": "月底",    "all_regions": False},
        {"id": "m6", "label": "儲值金預收",  "script": "儲值金預收.py",   "args": [],    "schedule": "月底",    "all_regions": False},
    ],
    "log_files": {},
}


def load_config():
    if CONFIG_F.exists():
        try:
            cfg = json.loads(CONFIG_F.read_text(encoding="utf-8"))
            if "log_files" not in cfg:
                cfg["log_files"] = {}
            return cfg
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
        "last_run": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
        "ok": ok,
        "stdout": (stdout or "")[-1500:],
        "stderr": (stderr or "")[-1500:],
    }
    save_runlog(log)


# ══════════════════════════════════════════════════
# Runner
# ══════════════════════════════════════════════════
REGION_ALL = "【全部地區依序】"


def run_script(script, args=None, region=None):
    if args is None:
        args = []
    path = BASE_DIR / script
    env  = os.environ.copy()
    if region and region in ACCOUNTS:
        acct = ACCOUNTS[region]
        env["REGION_NAME"]     = region
        env["REGION_EMAIL"]    = acct.get("email", "")
        env["REGION_PASSWORD"] = acct.get("password", "")
    cmd = [sys.executable, str(path)] + args
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True,
            timeout=300, cwd=str(BASE_DIR), env=env,
        )
        return {"ok": result.returncode == 0, "cmd": " ".join(cmd),
                "stdout": result.stdout, "stderr": result.stderr}
    except Exception as e:
        return {"ok": False, "cmd": " ".join(cmd), "stdout": "", "stderr": str(e)}


def do_run_job(job, region):
    targets = list(ACCOUNTS.keys()) if region == REGION_ALL else ([region] if region else [None])
    results = []
    for r in targets:
        res = run_script(job["script"], job.get("args", []), region=r)
        record_run(rkey(job["id"], r), res["ok"], res["stdout"], res["stderr"])
        results.append((r or "—", res))
    return results


# ══════════════════════════════════════════════════
# Output / Log helpers
# ══════════════════════════════════════════════════
def scan_output():
    files = []
    today_date = now_tw().date()
    for p in sorted(OUTPUT_DIR.rglob("*")):
        if p.is_file():
            s = p.stat()
            mdt = datetime.fromtimestamp(s.st_mtime, tz=TZ)
            files.append({
                "name":      p.name,
                "path":      p,
                "folder":    str(p.parent.relative_to(OUTPUT_DIR)) if p.parent != OUTPUT_DIR else "根目錄",
                "size_kb":   round(s.st_size / 1024, 1),
                "mtime":     mdt,
                "mtime_str": mdt.strftime("%Y-%m-%d %H:%M"),
                "today":     mdt.date() == today_date,
            })
    return sorted(files, key=lambda x: x["mtime"], reverse=True)


def read_last_lines(path, n=150):
    if isinstance(path, str):
        path = Path(path)
    if not path or not path.exists():
        return "(尚無 log 或檔案不存在)"
    try:
        lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
        return "\n".join(lines[-n:])
    except Exception as e:
        return f"(讀取失敗) {e}"


def highlight_log(text):
    html = []
    for line in text.splitlines():
        esc = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if any(k in line for k in ["Traceback", "Error", "ERROR", "❌", "PermissionError", "FAILED", "failed"]):
            html.append(f'<span class="le">{esc}</span>')
        elif any(k in line for k in ["✅", "SUCCESS", "success", "完成", "Done", "done"]):
            html.append(f'<span class="lo">{esc}</span>')
        elif any(k in line for k in ["WARNING", "Warning", "warn", "⚠"]):
            html.append(f'<span class="lw">{esc}</span>')
        elif any(k in line for k in ["INFO", "info", "開始", "Start", "start"]):
            html.append(f'<span class="li">{esc}</span>')
        else:
            html.append(f'<span class="ln">{esc}</span>')
    return "\n".join(html)


def calc_next_run(schedule_str: str) -> str:
    now = now_tw()
    try:
        s = schedule_str.strip()
        if len(s) == 5 and ":" in s:
            h, m = int(s[:2]), int(s[3:])
            candidate = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if candidate <= now:
                candidate += timedelta(days=1)
            return candidate.strftime("%Y-%m-%d  %H:%M")
        day_map = {"月初": 1, "月中": 15, "月底": 28}
        for prefix, day in day_map.items():
            if prefix in s:
                candidate = now.replace(day=day, hour=1, minute=0, second=0, microsecond=0)
                if candidate <= now:
                    if now.month == 12:
                        candidate = candidate.replace(year=now.year + 1, month=1)
                    else:
                        candidate = candidate.replace(month=now.month + 1)
                return candidate.strftime("%Y-%m-%d  %H:%M")
    except Exception:
        pass
    return "—"


def badge_html(cls, label):
    return f'<span class="badge {cls}"><span class="dot"></span>{label}</span>'


def file_size_str(path):
    if not path or not path.exists():
        return "—"
    s = path.stat().st_size
    if s < 1024:        return f"{s} B"
    if s < 1024*1024:   return f"{s/1024:.1f} KB"
    return f"{s/1024/1024:.1f} MB"


# ══════════════════════════════════════════════════
# Load data
# ══════════════════════════════════════════════════
cfg    = load_config()
runlog = load_runlog()
REGIONS = sorted(ACCOUNTS.keys()) if ACCOUNTS else []

# ══════════════════════════════════════════════════
# Top strip
# ══════════════════════════════════════════════════
now_str = now_tw().strftime("%Y/%m/%d  %H:%M")
rgn = st.session_state.region
chip_html = (
    f'<span class="rgn-chip">📍 {rgn}</span>'
    if rgn and rgn != REGION_ALL
    else '<span class="rgn-chip none">📍 未選地區</span>'
)
st.markdown(
    f'<div class="topstrip">'
    f'<div class="topstrip-brand">📊 營運報表控制台</div>'
    f'{chip_html}'
    f'<div class="topstrip-time">🕐 {now_str} 台北</div>'
    f'</div>',
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════
# Nav bar
# ══════════════════════════════════════════════════
PAGES = ["排程主控表", "手動執行", "Log 監控", "輸出報表", "腳本管理"]
ICONS = ["📋", "▶️", "📄", "📂", "⚙️"]
nav_cols = st.columns(len(PAGES) + 5)
for i, (pg, ic) in enumerate(zip(PAGES, ICONS)):
    active = st.session_state.page == pg
    wrap   = "nav-wrap active" if active else "nav-wrap"
    with nav_cols[i]:
        st.markdown(f'<div class="{wrap}">', unsafe_allow_html=True)
        if st.button(f"{ic} {pg}", key=f"nav_{pg}"):
            st.session_state.page = pg
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

page = st.session_state.page

# ══════════════════════════════════════════════════
# 地區選擇器
# ══════════════════════════════════════════════════
def region_selector(allow_all=False):
    if not ACCOUNTS:
        st.error("⚠️ 找不到 Streamlit secrets 帳密，請到 App settings > Secrets 設定。")
        return None
    opts = ["（不指定地區）"]
    if allow_all:
        opts.append(REGION_ALL)
    opts += REGIONS
    cur_rgn = st.session_state.region
    idx     = opts.index(cur_rgn) if cur_rgn in opts else 0
    c_sel, c_info = st.columns([2, 5])
    with c_sel:
        sel = st.selectbox("📍 操作地區", opts, index=idx, key="rgn_widget")
    with c_info:
        if sel not in ("（不指定地區）", REGION_ALL):
            if sel in ACCOUNTS:
                st.info(f"✅ **{sel}** 帳號：`{ACCOUNTS[sel].get('email','—')}`")
            else:
                st.warning(f"⚠️ 找不到「{sel}」帳密")
        elif sel == REGION_ALL:
            st.info(f"🌐 將依序執行所有地區：{', '.join(REGIONS)}")
        else:
            st.caption("未指定地區時腳本不會收到帳密環境變數")
    new_val = None if sel == "（不指定地區）" else sel
    if new_val != st.session_state.region:
        st.session_state.region = new_val
        st.rerun()
    return new_val


def show_run_result(label, pairs):
    st.markdown(f"**執行結果：{label}**")
    for rgn_name, res in pairs:
        prefix = f"`{rgn_name}` " if rgn_name != "—" else ""
        if res["ok"]:
            st.success(f"✅ {prefix}完成")
        else:
            st.error(f"❌ {prefix}失敗")
        with st.expander(f"{prefix}詳細輸出", expanded=not res["ok"]):
            st.code(res.get("cmd", ""), language="bash")
            st.text_area("stdout", res.get("stdout", "") or "(empty)", height=200,
                         key=f"_so_{label}_{rgn_name}_{id(res)}")
            st.text_area("stderr", res.get("stderr", "") or "(empty)", height=150,
                         key=f"_se_{label}_{rgn_name}_{id(res)}")


# ══════════════════════════════════════════════════
# PAGE: 排程主控表
# ══════════════════════════════════════════════════
if page == "排程主控表":
    st.markdown('<div class="pg-title">排程主控表</div><div class="pg-sub">SCHEDULE DASHBOARD</div>', unsafe_allow_html=True)
    region = region_selector()

    all_jobs   = cfg["daily"] + cfg["monthly"]
    ok_count   = sum(1 for j in all_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is True)
    fail_count = sum(1 for j in all_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False)
    wait_count = len(all_jobs) - ok_count - fail_count

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi blue"><div class="kpi-label">Total Tasks</div><div class="kpi-value">{len(all_jobs)}</div><div class="kpi-sub">已設定排程</div></div>
      <div class="kpi green"><div class="kpi-label">Success</div><div class="kpi-value">{ok_count}</div><div class="kpi-sub">執行成功</div></div>
      <div class="kpi red"><div class="kpi-label">Failed</div><div class="kpi-value">{fail_count}</div><div class="kpi-sub">執行失敗</div></div>
      <div class="kpi amber"><div class="kpi-label">Pending</div><div class="kpi-value">{wait_count}</div><div class="kpi-sub">待執行</div></div>
    </div>
    """, unsafe_allow_html=True)

    def render_task_section(jobs, title, section_key):
        st.markdown(f'<div class="panel"><div class="panel-head"><div class="panel-tag">{title}</div><div class="panel-note">勾選後可批次執行</div></div>', unsafe_allow_html=True)
        hcols = st.columns([0.4, 2.0, 1.4, 2.2, 1.8, 1.8, 0.9])
        for txt, col in zip(["", "任務名稱", "排程", "腳本", "狀態", "最後執行 / 下次預估", "執行"], hcols):
            col.markdown(f"<span style='font-family:\"DM Mono\",monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:#3a4f6a'>{txt}</span>", unsafe_allow_html=True)

        checked_jobs = []
        for job in jobs:
            k     = rkey(job["id"], region)
            entry = runlog.get(k) or runlog.get(job["id"], {})
            ok    = entry.get("ok", None)
            last  = entry.get("last_run", "—")
            last_short = last[:16] if last != "—" else "—"
            exists = (BASE_DIR / job["script"]).exists()

            if ok is True:
                st_badge = badge_html("green", "✓ 成功")
            elif ok is False:
                st_badge = badge_html("red",   "✗ 失敗")
            else:
                st_badge = badge_html("gray",  "— 待執行")

            e_icon = "🟢" if exists else "🔴"
            args_str = " ".join(job.get("args", []))
            script_display = f"{e_icon} `{job['script']}`" + (f" `{args_str}`" if args_str else "")
            next_run = calc_next_run(job.get("schedule", ""))

            row = st.columns([0.4, 2.0, 1.4, 2.2, 1.8, 1.8, 0.9])
            with row[0]:
                sel = st.checkbox("", key=f"batch_{section_key}_{job['id']}", label_visibility="collapsed")
                if sel:
                    checked_jobs.append(job)
            with row[1]:
                st.markdown(f"**{job['label']}**")
            with row[2]:
                st.caption(job.get("schedule", "—"))
            with row[3]:
                st.markdown(script_display)
            with row[4]:
                st.markdown(st_badge, unsafe_allow_html=True)
            with row[5]:
                st.caption(f"上次：{last_short}")
                if next_run != "—":
                    st.caption(f"↻ {next_run}")
            with row[6]:
                if st.button("▶", key=f"run_{section_key}_{job['id']}", help=f"執行 {job['label']}"):
                    with st.spinner(f"執行 {job['label']} …"):
                        pairs = do_run_job(job, region)
                    show_run_result(job["label"], pairs)
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)
        return checked_jobs

    checked_daily   = render_task_section(cfg["daily"],   "📅 每日排程", "daily")
    checked_monthly = render_task_section(cfg["monthly"], "🗓️ 月排程",   "monthly")

    all_checked = checked_daily + checked_monthly
    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">⚙️ 控制</div></div>', unsafe_allow_html=True)
    cc1, cc2, cc3 = st.columns([1, 1.5, 4])
    with cc1:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()
    with cc2:
        if st.button("▶ 執行已勾選", use_container_width=True):
            if not all_checked:
                st.warning("請先勾選至少一個任務")
            else:
                for job in all_checked:
                    with st.spinner(f"執行 {job['label']} …"):
                        pairs = do_run_job(job, region)
                    show_run_result(job["label"], pairs)
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    failures = [
        (j, runlog.get(rkey(j["id"], region)) or runlog.get(j["id"]))
        for j in all_jobs
        if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False
    ]
    if failures:
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">⚠️ 失敗詳情</div></div>', unsafe_allow_html=True)
        for job, entry in failures:
            with st.expander(f"❌ {job['label']}  ·  {entry.get('last_run','—')}"):
                if entry.get("stderr"):
                    st.code(entry["stderr"], language="bash")
                if entry.get("stdout"):
                    st.code(entry["stdout"])
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# PAGE: 手動執行
# ══════════════════════════════════════════════════
elif page == "手動執行":
    st.markdown('<div class="pg-title">手動執行</div><div class="pg-sub">MANUAL TRIGGER</div>', unsafe_allow_html=True)
    region = region_selector(allow_all=True)

    def run_and_show(job):
        with st.spinner(f"執行 {job['label']} …"):
            pairs = do_run_job(job, region)
        show_run_result(job["label"], pairs)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📅 每日報表</div></div>', unsafe_allow_html=True)
    cols = st.columns(5)
    for i, job in enumerate(cfg["daily"]):
        with cols[i % 5]:
            if st.button(job["label"], key=f"man_{job['id']}", use_container_width=True):
                run_and_show(job)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">🗓️ 月報表</div></div>', unsafe_allow_html=True)
    cols = st.columns(4)
    for i, job in enumerate(cfg["monthly"]):
        with cols[i % 4]:
            if st.button(job["label"], key=f"man_{job['id']}", use_container_width=True):
                run_and_show(job)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">🚀 批次執行</div></div>', unsafe_allow_html=True)
    bc1, bc2 = st.columns(2)

    def batch_run(jobs, title):
        results = []
        for job in jobs:
            with st.spinner(f"執行 {job['label']} …"):
                pairs = do_run_job(job, region)
            results.append((job["label"], pairs))
        st.markdown(f"**{title} 執行完畢**")
        for label, pairs in results:
            for rgn_name, res in pairs:
                prefix = f"`{rgn_name}` " if rgn_name != "—" else ""
                if res["ok"]:
                    st.success(f"✅ {prefix}{label}")
                else:
                    st.error(f"❌ {prefix}{label}")
                    if res.get("stderr"):
                        st.code(res["stderr"])

    with bc1:
        if st.button("▶ 全部每日報表", use_container_width=True):
            batch_run(cfg["daily"], "每日批次")
    with bc2:
        if st.button("▶ 全部月報表", use_container_width=True):
            batch_run(cfg["monthly"], "月報批次")
    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# PAGE: Log 監控
# ══════════════════════════════════════════════════
elif page == "Log 監控":
    st.markdown('<div class="pg-title">Log 監控</div><div class="pg-sub">LOG MONITOR</div>', unsafe_allow_html=True)

    builtin_logs = {"run_log（JSON 執行紀錄）": RUNLOG_F}
    if LOG_DIR.exists():
        for lf in sorted(LOG_DIR.glob("*.log")):
            builtin_logs[lf.name] = lf
    for name, path_str in cfg.get("log_files", {}).items():
        builtin_logs[name] = Path(path_str)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📄 Log 查看器</div></div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([3, 1, 1])
    with c1:
        sel_log_name = st.selectbox("選擇 Log 檔", list(builtin_logs.keys()), key="log_sel")
    with c2:
        n_lines = st.selectbox("顯示行數", [50, 100, 200, 500], index=1, key="log_lines")
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 刷新", use_container_width=True):
            st.rerun()

    log_path = builtin_logs[sel_log_name]

    if sel_log_name.startswith("run_log"):
        st.caption(f"檔案：{log_path}")
        if log_path.exists():
            try:
                import pandas as pd
                data = json.loads(log_path.read_text(encoding="utf-8"))
                rows = []
                for k, v in sorted(data.items(), key=lambda x: x[1].get("last_run", ""), reverse=True):
                    rows.append({
                        "key":            k,
                        "last_run":       v.get("last_run", "—"),
                        "狀態":           "✅ 成功" if v.get("ok") else "❌ 失敗",
                        "stderr_preview": (v.get("stderr") or "")[:80],
                    })
                st.dataframe(pd.DataFrame(rows)[["key", "last_run", "狀態", "stderr_preview"]],
                             use_container_width=True, hide_index=True)
                with st.expander("查看完整 JSON"):
                    st.json(data)
            except Exception as e:
                st.error(f"無法解析 JSON：{e}")
        else:
            st.info("尚無執行紀錄")
    else:
        mtime_str = ""
        if log_path.exists():
            mtime_str = datetime.fromtimestamp(log_path.stat().st_mtime, tz=TZ).strftime("%Y-%m-%d %H:%M:%S")
        st.caption(f"檔案：{log_path}  ·  更新：{mtime_str}")
        raw = read_last_lines(log_path, n_lines)
        st.markdown(f'<div class="logbox">{highlight_log(raw)}</div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">➕ 新增監控 Log 路徑</div></div>', unsafe_allow_html=True)
    with st.form("add_log_form"):
        l1, l2 = st.columns(2)
        with l1:
            log_alias = st.text_input("別名", placeholder="例：cron.log")
        with l2:
            log_fpath = st.text_input("完整路徑", placeholder="/absolute/path/to/file.log")
        if st.form_submit_button("➕ 加入"):
            if log_alias and log_fpath:
                cfg["log_files"][log_alias] = log_fpath
                save_config(cfg)
                st.success(f"已加入：{log_alias}")
                st.rerun()
            else:
                st.error("別名與路徑為必填")

    extra = cfg.get("log_files", {})
    if extra:
        st.markdown("**已設定的額外 Log：**")
        for name in list(extra.keys()):
            rc1, rc2 = st.columns([4, 1])
            with rc1:
                st.caption(f"`{name}` → {extra[name]}")
            with rc2:
                if st.button("移除", key=f"rm_log_{name}"):
                    del cfg["log_files"][name]
                    save_config(cfg)
                    st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# PAGE: 輸出報表
# ══════════════════════════════════════════════════
elif page == "輸出報表":
    st.markdown('<div class="pg-title">輸出報表</div><div class="pg-sub">OUTPUT FILES</div>', unsafe_allow_html=True)

    files       = scan_output()
    folders     = sorted({f["folder"] for f in files})
    today_files = [f for f in files if f["today"]]

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi blue"><div class="kpi-label">Total Files</div><div class="kpi-value">{len(files)}</div><div class="kpi-sub">output/ 下所有檔案</div></div>
      <div class="kpi green"><div class="kpi-label">Today</div><div class="kpi-value">{len(today_files)}</div><div class="kpi-sub">今日產出</div></div>
      <div class="kpi amber"><div class="kpi-label">Total Size</div><div class="kpi-value">{sum(f['size_kb'] for f in files):.0f}</div><div class="kpi-sub">KB</div></div>
      <div class="kpi blue"><div class="kpi-label">Folders</div><div class="kpi-value">{len(folders)}</div><div class="kpi-sub">子目錄數</div></div>
    </div>
    """, unsafe_allow_html=True)

    if today_files:
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">🟢 今日產出</div></div>', unsafe_allow_html=True)
        hcols = st.columns([4, 2, 2, 1])
        for txt, col in zip(["檔名", "資料夾", "時間", "大小"], hcols):
            col.markdown(f"<span style='font-family:\"DM Mono\",monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:#3a4f6a'>{txt}</span>", unsafe_allow_html=True)
        st.divider()
        for f in today_files[:20]:
            row = st.columns([4, 2, 2, 1])
            row[0].markdown(f"📄 **{f['name']}**")
            row[1].caption(f["folder"])
            row[2].caption(f["mtime_str"])
            row[3].caption(f"{f['size_kb']} KB")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📁 全部檔案</div></div>', unsafe_allow_html=True)
    c_fld, c_kw, c_ref = st.columns([2, 3, 1])
    with c_fld:
        sel_folder = st.selectbox("篩選資料夾", ["（全部）"] + folders, key="out_folder")
    with c_kw:
        kw = st.text_input("搜尋檔名", placeholder="輸入關鍵字…", key="out_kw")
    with c_ref:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 重新掃描", use_container_width=True):
            st.rerun()

    filtered = files
    if sel_folder != "（全部）":
        filtered = [f for f in filtered if f["folder"] == sel_folder]
    if kw:
        filtered = [f for f in filtered if kw.lower() in f["name"].lower()]

    if not filtered:
        st.info("output/ 目前沒有符合條件的檔案")
    else:
        hcols = st.columns([4, 2, 2, 1, 1])
        for txt, col in zip(["檔名", "資料夾", "修改時間", "大小", "今日"], hcols):
            col.markdown(f"<span style='font-family:\"DM Mono\",monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:#3a4f6a'>{txt}</span>", unsafe_allow_html=True)
        st.divider()
        for f in filtered:
            row = st.columns([4, 2, 2, 1, 1])
            row[0].markdown(f"📄 **{f['name']}**")
            row[1].caption(f["folder"])
            row[2].caption(f["mtime_str"])
            row[3].caption(f"{f['size_kb']} KB")
            row[4].markdown(
                badge_html("green", "今日") if f["today"] else badge_html("gray", "—"),
                unsafe_allow_html=True,
            )
    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# PAGE: 腳本管理
# ══════════════════════════════════════════════════
elif page == "腳本管理":
    st.markdown('<div class="pg-title">腳本管理</div><div class="pg-sub">SCRIPT MANAGEMENT</div>', unsafe_allow_html=True)

    tab_d, tab_m, tab_edit, tab_reset = st.tabs(["📅 每日腳本", "🗓️ 月腳本", "📝 線上編輯", "🔧 重置"])

    def new_id(prefix="x"):
        return f"{prefix}{now_tw().strftime('%f')}"

    def job_editor(group_key, tab):
        with tab:
            jobs = cfg[group_key]
            st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">現有腳本</div></div>', unsafe_allow_html=True)
            for idx, job in enumerate(jobs):
                exists = (BASE_DIR / job["script"]).exists()
                icon   = "🟢" if exists else "🔴"
                with st.expander(f"{icon} {job['label']}  ·  `{job['script']}`"):
                    c1, c2 = st.columns(2)
                    with c1:
                        new_label  = st.text_input("顯示名稱",           job["label"],                  key=f"lbl_{job['id']}")
                        new_script = st.text_input("腳本檔名",            job["script"],                 key=f"scr_{job['id']}")
                    with c2:
                        new_args   = st.text_input("參數（空格分隔）",    " ".join(job.get("args", [])), key=f"arg_{job['id']}")
                        new_sched  = st.text_input("排程說明",            job.get("schedule", ""),       key=f"sch_{job['id']}")
                    new_multi = st.checkbox("需跑全部地區", value=job.get("all_regions", False), key=f"mul_{job['id']}")
                    next_run = calc_next_run(new_sched)
                    if next_run != "—":
                        st.markdown(f'<div class="next-run">⏭️ 下次執行預估：<strong>{next_run}</strong></div>', unsafe_allow_html=True)
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
                        if st.button("▶ 測試", key=f"test_{job['id']}"):
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

            st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">➕ 新增腳本</div></div>', unsafe_allow_html=True)
            with st.form(key=f"add_{group_key}"):
                c1, c2 = st.columns(2)
                with c1:
                    f_label  = st.text_input("顯示名稱",         placeholder="例：新報表")
                    f_script = st.text_input("腳本檔名",         placeholder="例：新報表.py")
                with c2:
                    f_args   = st.text_input("參數（空格分隔）",  placeholder="例：0800")
                    f_sched  = st.text_input("排程時間說明",      placeholder="例：09:00 或 月初01日")
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
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📝 直接編輯腳本內容</div></div>', unsafe_allow_html=True)
        all_scripts = sorted({j["script"] for grp in ["daily", "monthly"] for j in cfg[grp]})
        sel = st.selectbox("選擇腳本", ["（請選擇）"] + all_scripts, key="edit_sel")
        if sel and sel != "（請選擇）":
            spath = BASE_DIR / sel
            code  = spath.read_text(encoding="utf-8") if spath.exists() else f"# {sel} 尚未建立\n"
            edited = st.text_area("腳本內容", code, height=460, key="edit_content")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("💾 寫入儲存", use_container_width=True):
                    spath.write_text(edited, encoding="utf-8")
                    st.success(f"已寫入 {sel}")
            with c2:
                st.caption(f"路徑：{spath}")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_reset:
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">⚠️ 重置設定</div></div>', unsafe_allow_html=True)
        st.warning("重置後所有排程設定將還原為預設值，run_log 不受影響。")
        confirm = st.checkbox("我確認要重置所有排程設定")
        if st.button("🔄 確認重置", disabled=not confirm):
            save_config(DEFAULT_CONFIG)
            st.warning("已重置，請重新整理頁面")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# Footer
# ══════════════════════════════════════════════════
st.markdown("<br>", unsafe_allow_html=True)
st.caption(
    f"營運報表控制台 v5.1  ·  {now_tw().strftime('%Y-%m-%d %H:%M:%S')} 台北  ·  "
    f"可用地區：{', '.join(REGIONS) if REGIONS else '尚未設定'}"
)
