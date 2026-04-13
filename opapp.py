"""
opapp.py  ──  營運報表控制台 v6.1
延續 v6 架構，新增：
  - CLI 排程入口：python opapp.py --run-scheduled
  - 支援 daily / monthly 自動判斷
  - 支援 schedule_config.json 覆蓋 DEFAULT_CONFIG
  - 台北時區判斷
  - 同日已成功執行則略過，避免重複跑
"""

import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st

TZ = ZoneInfo("Asia/Taipei")
REGION_ALL = "【全部地區依序】"


def now_tw():
    return datetime.now(TZ)


# =========================================================
# CLI mode detection
# =========================================================
CLI_ARGS = set(sys.argv[1:])
IS_CLI_MODE = (
    "--run-scheduled" in CLI_ARGS
    or "--list-jobs" in CLI_ARGS
    or "--run-job" in CLI_ARGS
)


# =========================================================
# 帳密
# =========================================================
def load_accounts():
    try:
        acc_root = st.secrets["accounts"]
    except Exception:
        return {}

    mapping = {
        "台北": "taipei",
        "台中": "taichung",
        "桃園": "taoyuan",
        "新竹": "hsinchu",
        "高雄": "kaohsiung",
    }

    out = {}
    for region_name, secret_key in mapping.items():
        try:
            item = acc_root[secret_key]
            email = item.get("email", "")
            password = item.get("password", "")
            if email and password:
                out[region_name] = {
                    "email": email,
                    "password": password,
                }
        except Exception:
            continue
    return out


ACCOUNTS = load_accounts()


# =========================================================
# Paths
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"
CONFIG_F = BASE_DIR / "schedule_config.json"
RUNLOG_F = BASE_DIR / "run_log.json"
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)


# =========================================================
# Page config
# =========================================================
if not IS_CLI_MODE:
    st.set_page_config(
        page_title="營運報表控制台",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&family=Noto+Sans+TC:wght@400;500;700&display=swap');

    :root {
        --bg:       #eef1f6;
        --surface:  #ffffff;
        --surface2: #f4f6f9;
        --border:   #dde2ea;
        --border2:  #c8cfd9;
        --text-1:   #0f172a;
        --text-2:   #1e293b;
        --text-3:   #475569;
        --text-4:   #94a3b8;
        --accent:   #0d9488;
        --accent-dk:#0b7c73;
        --accent-lt:#f0fdfa;
        --accent-bd:#99f6e4;
        --blue:     #1e40af;
        --blue-lt:  #eff6ff;
        --blue-bd:  #bfdbfe;
        --green:    #166534;
        --green-lt: #f0fdf4;
        --green-bd: #bbf7d0;
        --red:      #991b1b;
        --red-lt:   #fff1f2;
        --red-bd:   #fecaca;
        --amber:    #92400e;
        --amber-lt: #fffbeb;
        --amber-bd: #fde68a;
        --radius:   10px;
        --shadow-sm: 0 1px 2px rgba(0,0,0,.06);
        --shadow:    0 1px 3px rgba(0,0,0,.08), 0 4px 12px rgba(0,0,0,.05);
    }

    html, body, [class*="css"] {
        font-family: 'Noto Sans TC', 'Outfit', sans-serif !important;
        background: var(--bg) !important;
        color: var(--text-1) !important;
    }
    #MainMenu, footer, header { visibility: hidden; }
    section[data-testid="stSidebar"],
    [data-testid="collapsedControl"] { display: none !important; }
    .block-container { padding-top: 0 !important; padding-bottom: 3rem; max-width: 1380px; }

    .topbar {
        background: var(--surface);
        border-bottom: 1.5px solid var(--border);
        padding: 0 32px;
        margin: 0 -1rem;
        display: flex; align-items: center;
        height: 56px;
        position: sticky; top: 0; z-index: 9999;
        box-shadow: var(--shadow);
        gap: 16px;
    }
    .topbar-brand {
        font-family: 'Outfit', sans-serif;
        font-size: 16px; font-weight: 800;
        color: var(--accent);
        letter-spacing: -.02em;
        white-space: nowrap; flex-shrink: 0;
    }
    .topbar-sep { width:1px; height:20px; background:var(--border); flex-shrink:0; }
    .topbar-time {
        margin-left: auto;
        font-family: 'JetBrains Mono', monospace;
        font-size: 12px; color: var(--text-3); flex-shrink: 0;
    }
    .rgn-chip {
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px; font-weight: 500;
        background: var(--accent-lt);
        border: 1px solid var(--accent-bd);
        color: var(--accent);
        padding: 3px 12px; border-radius: 20px;
    }
    .rgn-chip.none {
        background: var(--surface2); border-color: var(--border); color: var(--text-4);
    }

    div[data-testid="stHorizontalBlock"]:has(div.nav-wrap) {
        background: var(--surface) !important;
        border-bottom: 1.5px solid var(--border) !important;
        padding: 0 24px !important;
        margin: 0 -1rem 28px !important;
        gap: 0 !important;
    }
    .nav-wrap div[data-testid="stButton"] > button {
        height: 46px !important; padding: 0 20px !important;
        background: transparent !important; border: none !important;
        border-bottom: 2px solid transparent !important; border-radius: 0 !important;
        color: var(--text-3) !important;
        font-family: 'Outfit', sans-serif !important; font-weight: 700 !important; font-size: 13.5px !important;
        box-shadow: none !important; transition: color .15s, border-color .15s !important;
    }
    .nav-wrap div[data-testid="stButton"] > button:hover { color: var(--accent) !important; border-bottom-color: var(--accent-bd) !important; }
    .nav-wrap.active div[data-testid="stButton"] > button { color: var(--accent) !important; border-bottom-color: var(--accent) !important; }

    .pg-title { font-family:'Outfit',sans-serif; font-size:22px; font-weight:800; color:var(--text-1); margin-bottom:2px; letter-spacing:-.02em; }
    .pg-sub   { font-family:'JetBrains Mono',monospace; font-size:10.5px; color:var(--text-4); letter-spacing:.14em; margin-bottom:22px; }

    .kpi-row { display:flex; gap:14px; margin-bottom:24px; }
    .kpi {
        flex:1; background:var(--surface); border:1px solid var(--border);
        border-radius:14px; padding:18px 22px; position:relative; overflow:hidden;
        box-shadow: var(--shadow);
    }
    .kpi::after { content:''; position:absolute; top:0; left:0; right:0; height:3px; }
    .kpi.blue::after  { background: linear-gradient(90deg,#1e40af,#60a5fa); }
    .kpi.green::after { background: linear-gradient(90deg,#166534,#4ade80); }
    .kpi.red::after   { background: linear-gradient(90deg,#991b1b,#f87171); }
    .kpi.amber::after { background: linear-gradient(90deg,#92400e,#fbbf24); }
    .kpi-label { font-family:'JetBrains Mono',monospace; font-size:10px; font-weight:500; letter-spacing:.14em; text-transform:uppercase; color:var(--text-4); margin-bottom:8px; }
    .kpi-value { font-family:'Outfit',sans-serif; font-size:32px; font-weight:800; color:var(--text-1); line-height:1; }
    .kpi-sub   { font-size:12px; color:var(--text-3); margin-top:4px; }

    .panel { background:var(--surface); border:1px solid var(--border); border-radius:14px; padding:20px 24px 18px; margin-bottom:18px; box-shadow:var(--shadow); }
    .panel-head { display:flex; align-items:center; gap:10px; margin-bottom:16px; padding-bottom:13px; border-bottom:1px solid var(--border); }
    .panel-tag { font-family:'JetBrains Mono',monospace; font-size:10.5px; font-weight:500; letter-spacing:.12em; text-transform:uppercase; color:var(--accent); background:var(--accent-lt); border:1px solid var(--accent-bd); border-radius:6px; padding:3px 10px; }
    .panel-note { font-size:12px; color:var(--text-4); margin-left:auto; }

    .th { font-family:'JetBrains Mono',monospace; font-size:10px; font-weight:500; letter-spacing:.12em; text-transform:uppercase; color:var(--text-4); padding-bottom:6px; border-bottom:1px solid var(--border); }

    .sched-lbl {
        display: inline-flex; align-items: center; gap: 5px;
        font-family: 'JetBrains Mono', monospace; font-size: 12px; font-weight: 500;
        color: var(--text-2); background: var(--surface2);
        border: 1px solid var(--border2); border-radius: 7px;
        padding: 4px 10px; cursor: default;
    }
    .next-time { font-family:'JetBrains Mono',monospace; font-size:11px; color:var(--text-4); margin-top:2px; }

    .badge { display:inline-flex; align-items:center; gap:5px; font-family:'JetBrains Mono',monospace; font-size:11px; font-weight:500; padding:3px 10px; border-radius:20px; white-space:nowrap; }
    .badge .dot { width:6px; height:6px; border-radius:50%; flex-shrink:0; }
    .badge.green { background:var(--green-lt); border:1px solid var(--green-bd); color:var(--green); }
    .badge.green .dot  { background:#22c55e; }
    .badge.red   { background:var(--red-lt);   border:1px solid var(--red-bd);   color:var(--red); }
    .badge.red   .dot  { background:#ef4444; }
    .badge.amber { background:var(--amber-lt); border:1px solid var(--amber-bd); color:var(--amber); }
    .badge.amber .dot  { background:#f59e0b; }
    .badge.gray  { background:var(--surface2); border:1px solid var(--border2);  color:var(--text-3); }
    .badge.gray  .dot  { background:#94a3b8; }
    .badge.teal  { background:var(--accent-lt); border:1px solid var(--accent-bd); color:var(--accent); }
    .badge.teal  .dot  { background:var(--accent); }
    .badge.blue  { background:var(--blue-lt); border:1px solid var(--blue-bd);   color:var(--blue); }
    .badge.blue  .dot  { background:#60a5fa; }

    .logbox {
        background: #1a2435; border:1px solid #2a3a52; border-radius:12px;
        padding:16px 20px; font-family:'JetBrains Mono',monospace; font-size:12.5px; line-height:1.75;
        white-space:pre-wrap; word-break:break-all; max-height:520px; overflow-y:auto;
    }
    .le { color:#fca5a5; display:block; } .lo { color:#86efac; display:block; }
    .lw { color:#fde68a; display:block; } .li { color:#93c5fd; display:block; }
    .ln { color:#8fa3b8; display:block; }

    .next-run { font-family:'JetBrains Mono',monospace; font-size:12px; color:var(--accent); background:var(--accent-lt); border:1px solid var(--accent-bd); border-radius:8px; padding:8px 14px; margin-top:8px; display:flex; align-items:center; gap:8px; }

    .stButton > button {
        background: var(--surface) !important; color: var(--text-1) !important;
        border: 1.5px solid var(--border2) !important; border-radius: 8px !important;
        font-family: 'Noto Sans TC', sans-serif !important; font-weight: 600 !important; font-size: 13px !important;
        transition: all .15s !important;
    }
    .stButton > button:hover { background: var(--accent-lt) !important; color: var(--accent) !important; border-color: var(--accent-bd) !important; }
    .stButton > button[kind="primary"] { background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important; }
    .stButton > button[kind="primary"]:hover { background: var(--accent-dk) !important; color: #fff !important; }

    .stSelectbox label, .stTextInput label, .stTextArea label,
    .stCheckbox label, .stMultiSelect label {
        color: var(--text-2) !important; font-size: 13px !important; font-weight: 600 !important;
    }
    .stSelectbox > div > div { background:var(--surface) !important; color:var(--text-1) !important; border:1.5px solid var(--border2) !important; border-radius:8px !important; font-size:13px !important; }
    .stTextInput input, .stTextArea textarea { background:var(--surface) !important; color:var(--text-1) !important; border:1.5px solid var(--border2) !important; border-radius:8px !important; font-size:13px !important; }
    .stTextInput input:focus, .stTextArea textarea:focus { border-color:var(--accent) !important; box-shadow:0 0 0 3px rgba(13,148,136,.12) !important; outline: none !important; }

    [data-testid="stMetricLabel"]  { color: var(--text-3) !important; font-size: 12px !important; }
    [data-testid="stMetricValue"]  { color: var(--text-1) !important; font-size: 26px !important; font-weight: 800 !important; }
    [data-testid="metric-container"] { background:var(--surface) !important; border:1px solid var(--border) !important; border-radius:12px !important; padding:16px 20px !important; box-shadow:var(--shadow) !important; }

    .stTabs [data-baseweb="tab-list"] { background:transparent !important; border-bottom:1px solid var(--border) !important; gap:0 !important; }
    .stTabs [data-baseweb="tab"] { background:transparent !important; color:var(--text-3) !important; font-size:13px !important; font-weight:700 !important; padding:8px 20px !important; border-radius:0 !important; }
    .stTabs [aria-selected="true"] { color:var(--accent) !important; border-bottom:2px solid var(--accent) !important; }

    .stCaption, [data-testid="stCaptionContainer"] { color: var(--text-3) !important; font-size: 12px !important; }
    details > summary, .streamlit-expanderHeader { color: var(--text-2) !important; font-size: 13px !important; font-weight: 600 !important; background: var(--surface2) !important; }
    [data-testid="stExpander"] { border: 1px solid var(--border) !important; border-radius: var(--radius) !important; }

    hr { border-color: var(--border) !important; margin: 6px 0 14px !important; }
    [data-testid="stCodeBlock"] { background: #1a2435 !important; border: 1px solid #2a3a52 !important; border-radius: 8px !important; }
    [data-testid="stCodeBlock"] code, [data-testid="stCodeBlock"] * { color: #e2e8f0 !important; }
    [data-testid="stAlert"] { border-radius: 8px !important; }

    [data-testid="stDataFrame"] { border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow-sm); }
    </style>
    """, unsafe_allow_html=True)


# =========================================================
# Session state
# =========================================================
if not IS_CLI_MODE:
    for k, v in {"page": "排程主控表", "region": None}.items():
        if k not in st.session_state:
            st.session_state[k] = v


# =========================================================
# Config / RunLog
# =========================================================
DEFAULT_CONFIG = {
    "daily": [
        {"id": "d1", "label": "排班統計表",   "script": "schedule_report.py",       "args": [],    "schedule": "01:10",       "all_regions": False, "enabled": True},
        {"id": "d2", "label": "專員班表",     "script": "staff_schedule.py",        "args": [],    "schedule": "01:20",       "all_regions": False, "enabled": True},
        {"id": "d3", "label": "專員個資",     "script": "staff_info.py",            "args": [],    "schedule": "01:30",       "all_regions": False, "enabled": True},
        {"id": "d4", "label": "當月次月訂單", "script": "orders_report.py",         "args": [],    "schedule": "01:40",       "all_regions": False, "enabled": True},
        {"id": "d5", "label": "業績報表",     "script": "performance_report.py",    "args": [],    "schedule": "08:00",       "all_regions": True, "enabled": True},
    ],
    "monthly": [
        {"id": "m1", "label": "上半月訂單",   "script": "half_month_orders.py",       "args": ["1"], "schedule": "每月15日18:15", "all_regions": True, "enabled": True},
        {"id": "m2", "label": "下半月訂單",   "script": "half_month_orders.py",       "args": ["2"], "schedule": "每月底18:15",  "all_regions": True, "enabled": True},
        {"id": "m3", "label": "已退款",       "script": "refund_report.py",           "args": [],    "schedule": "月底18:30",    "all_regions": True, "enabled": True},
        {"id": "m4", "label": "預收",         "script": "prepaid_report.py",          "args": [],    "schedule": "月初00:10",    "all_regions": True, "enabled": True},
        {"id": "m5", "label": "儲值金結算",   "script": "stored_value_settlement.py", "args": [],    "schedule": "月初00:20",    "all_regions": True, "enabled": True},
        {"id": "m6", "label": "儲值金預收",   "script": "stored_value_prepaid.py",    "args": [],    "schedule": "月初00:30",    "all_regions": True, "enabled": True},
    ],
    "log_files": {},
}


def _deepcopy_config(obj):
    return json.loads(json.dumps(obj, ensure_ascii=False))


def load_config():
    if CONFIG_F.exists():
        try:
            cfg = json.loads(CONFIG_F.read_text(encoding="utf-8"))
            cfg.setdefault("daily", [])
            cfg.setdefault("monthly", [])
            cfg.setdefault("log_files", {})
            for grp in ["daily", "monthly"]:
                for job in cfg.get(grp, []):
                    job.setdefault("args", [])
                    job.setdefault("schedule", "")
                    job.setdefault("all_regions", False)
                    job.setdefault("enabled", True)
            return cfg
        except Exception:
            pass
    return _deepcopy_config(DEFAULT_CONFIG)


def save_config(cfg):
    CONFIG_F.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_runlog():
    if RUNLOG_F.exists():
        try:
            return json.loads(RUNLOG_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_runlog(log):
    RUNLOG_F.write_text(
        json.dumps(log, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def rkey(job_id, region=None):
    return f"{region}__{job_id}" if region else job_id


def record_run(key, ok, stdout, stderr):
    log = load_runlog()
    log[key] = {
        "last_run": now_tw().strftime("%Y-%m-%d %H:%M:%S"),
        "ok": ok,
        "stdout": (stdout or "")[-3000:],
        "stderr": (stderr or "")[-3000:],
    }
    save_runlog(log)


# =========================================================
# Runner
# =========================================================
def run_script(script, args=None, region=None):
    path = BASE_DIR / script
    env = os.environ.copy()

    if region and region in ACCOUNTS:
        acct = ACCOUNTS[region]
        env["REGION_NAME"] = region
        env["REGION_EMAIL"] = acct.get("email", "")
        env["REGION_PASSWORD"] = acct.get("password", "")

    cmd = [sys.executable, str(path)] + [str(a) for a in (args or [])]

    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=900,
            cwd=str(BASE_DIR),
            env=env,
        )
        return {
            "ok": r.returncode == 0,
            "cmd": " ".join(cmd),
            "stdout": r.stdout,
            "stderr": r.stderr,
        }
    except Exception as e:
        return {
            "ok": False,
            "cmd": " ".join(cmd),
            "stdout": "",
            "stderr": str(e),
        }


def do_run_job(job, region, extra_args=None):
    args = job.get("args", []) + (extra_args or [])

    if job.get("all_regions", False):
        targets = list(ACCOUNTS.keys())
    else:
        targets = [region] if region else [None]

    results = []
    for r in targets:
        res = run_script(job["script"], args, region=r)
        record_run(rkey(job["id"], r), res["ok"], res["stdout"], res["stderr"])
        results.append((r or "—", res))
    return results


# =========================================================
# Schedule parser / auto runner
# =========================================================
def is_last_day(dt: datetime) -> bool:
    return (dt + timedelta(days=1)).month != dt.month


def parse_schedule_spec(schedule_str: str):
    """
    回傳 dict:
    {
        "type": "daily" | "monthly",
        "day_kind": "fixed" | "start" | "mid" | "end" | None,
        "day": int | None,
        "hour": int,
        "minute": int,
        "raw": str
    }

    支援：
    - 01:10
    - 每月15日18:15
    - 每月底18:15
    - 月底18:30
    - 月初00:10
    - 月中18:15
    - 月初01日  -> 01日 00:00
    - 月底28日  -> 28日 00:00
    """
    s = (schedule_str or "").strip()
    if not s:
        return None

    # daily: HH:MM
    if re.fullmatch(r"\d{2}:\d{2}", s):
        hour, minute = map(int, s.split(":"))
        return {
            "type": "daily",
            "day_kind": None,
            "day": None,
            "hour": hour,
            "minute": minute,
            "raw": s,
        }

    # extract time if present
    tm = re.search(r"(\d{1,2}):(\d{2})", s)
    if tm:
        hour = int(tm.group(1))
        minute = int(tm.group(2))
    else:
        hour = 0
        minute = 0

    if "月底" in s or "每月底" in s:
        return {
            "type": "monthly",
            "day_kind": "end",
            "day": None,
            "hour": hour,
            "minute": minute,
            "raw": s,
        }

    if "月初" in s:
        day_match = re.search(r"月初(\d{1,2})日", s)
        day = int(day_match.group(1)) if day_match else 1
        return {
            "type": "monthly",
            "day_kind": "fixed",
            "day": day,
            "hour": hour,
            "minute": minute,
            "raw": s,
        }

    if "月中" in s:
        day_match = re.search(r"月中(\d{1,2})日", s)
        day = int(day_match.group(1)) if day_match else 15
        return {
            "type": "monthly",
            "day_kind": "fixed",
            "day": day,
            "hour": hour,
            "minute": minute,
            "raw": s,
        }

    fixed = re.search(r"(?:每月)?(\d{1,2})日", s)
    if fixed:
        return {
            "type": "monthly",
            "day_kind": "fixed",
            "day": int(fixed.group(1)),
            "hour": hour,
            "minute": minute,
            "raw": s,
        }

    return None


def schedule_matches_now(schedule_str: str, now: datetime | None = None) -> bool:
    now = now or now_tw()
    spec = parse_schedule_spec(schedule_str)
    if not spec:
        return False

    if now.hour != spec["hour"] or now.minute != spec["minute"]:
        return False

    if spec["type"] == "daily":
        return True

    if spec["day_kind"] == "end":
        return is_last_day(now)

    if spec["day_kind"] == "fixed":
        return now.day == spec["day"]

    return False


def calc_next_run(schedule_str: str) -> str:
    now = now_tw()
    spec = parse_schedule_spec(schedule_str)
    if not spec:
        return "—"

    try:
        if spec["type"] == "daily":
            candidate = now.replace(
                hour=spec["hour"],
                minute=spec["minute"],
                second=0,
                microsecond=0,
            )
            if candidate <= now:
                candidate += timedelta(days=1)
            return candidate.strftime("%m/%d  %H:%M")

        # monthly
        def month_shift(dt: datetime):
            if dt.month == 12:
                return dt.replace(year=dt.year + 1, month=1, day=1)
            return dt.replace(month=dt.month + 1, day=1)

        base = now.replace(second=0, microsecond=0)

        if spec["day_kind"] == "end":
            probe = base
            while not is_last_day(probe):
                probe += timedelta(days=1)
            candidate = probe.replace(
                hour=spec["hour"],
                minute=spec["minute"],
                second=0,
                microsecond=0,
            )
            if candidate <= now:
                nxt = month_shift(now)
                probe = nxt
                while not is_last_day(probe):
                    probe += timedelta(days=1)
                candidate = probe.replace(
                    hour=spec["hour"],
                    minute=spec["minute"],
                    second=0,
                    microsecond=0,
                )
            return candidate.strftime("%m/%d  %H:%M")

        if spec["day_kind"] == "fixed":
            day = spec["day"]
            try:
                candidate = now.replace(
                    day=day,
                    hour=spec["hour"],
                    minute=spec["minute"],
                    second=0,
                    microsecond=0,
                )
            except ValueError:
                candidate = now.replace(
                    day=28,
                    hour=spec["hour"],
                    minute=spec["minute"],
                    second=0,
                    microsecond=0,
                )

            if candidate <= now:
                nxt = month_shift(now)
                try:
                    candidate = nxt.replace(
                        day=day,
                        hour=spec["hour"],
                        minute=spec["minute"],
                        second=0,
                        microsecond=0,
                    )
                except ValueError:
                    candidate = nxt.replace(
                        day=28,
                        hour=spec["hour"],
                        minute=spec["minute"],
                        second=0,
                        microsecond=0,
                    )
            return candidate.strftime("%m/%d  %H:%M")
    except Exception:
        pass

    return "—"


def has_success_run_today(job_id: str, region=None) -> bool:
    runlog = load_runlog()
    entry = runlog.get(rkey(job_id, region)) or runlog.get(job_id) or {}
    if entry.get("ok") is not True:
        return False
    last_run = entry.get("last_run", "")
    try:
        last_dt = datetime.strptime(last_run, "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ)
        return last_dt.date() == now_tw().date()
    except Exception:
        return False


def run_scheduled_jobs():
    cfg = load_config()
    now = now_tw()
    all_jobs = cfg.get("daily", []) + cfg.get("monthly", [])

    print(f"[INFO] Taipei now: {now.strftime('%Y-%m-%d %H:%M:%S')}")

    matched = []
    for job in all_jobs:
        if not job.get("enabled", True):
            continue
        if schedule_matches_now(job.get("schedule", ""), now):
            matched.append(job)

    if not matched:
        print("[INFO] No scheduled jobs matched this minute.")
        return 0

    exit_code = 0

    for job in matched:
        print(f"[INFO] Matched job: {job['id']} / {job['label']} / {job.get('schedule','')}")

        if job.get("all_regions", False):
            targets = list(ACCOUNTS.keys())
            if not targets:
                print(f"[WARN] {job['label']} 需要 all_regions，但目前沒有可用帳號")
                exit_code = 1
                continue
        else:
            targets = [None]

        for region in targets:
            if has_success_run_today(job["id"], region):
                print(f"[INFO] Skip already succeeded today: {job['label']} / {region or '—'}")
                continue

            print(f"[INFO] Running {job['label']} / {region or '—'}")
            res = run_script(job["script"], job.get("args", []), region=region)
            record_run(rkey(job["id"], region), res["ok"], res["stdout"], res["stderr"])

            if res["ok"]:
                print(f"[OK] {job['label']} / {region or '—'}")
            else:
                print(f"[ERROR] {job['label']} / {region or '—'}")
                if res.get("stderr"):
                    print(res["stderr"][-1000:])
                exit_code = 1

    return exit_code


def list_jobs_cli():
    cfg = load_config()
    for grp in ["daily", "monthly"]:
        print(f"\n[{grp.upper()}]")
        for job in cfg.get(grp, []):
            print(
                f"- {job['id']}: {job['label']} | {job['script']} | "
                f"schedule={job.get('schedule','')} | "
                f"enabled={job.get('enabled', True)} | "
                f"all_regions={job.get('all_regions', False)}"
            )


def run_one_job_cli(job_id: str):
    cfg = load_config()
    all_jobs = cfg.get("daily", []) + cfg.get("monthly", [])
    target = next((j for j in all_jobs if j.get("id") == job_id), None)

    if not target:
        print(f"[ERROR] job id not found: {job_id}")
        return 1

    print(f"[INFO] Run single job: {target['label']}")

    if target.get("all_regions", False):
        targets = list(ACCOUNTS.keys())
        if not targets:
            print("[ERROR] all_regions job but no accounts configured")
            return 1
    else:
        targets = [None]

    exit_code = 0
    for region in targets:
        res = run_script(target["script"], target.get("args", []), region=region)
        record_run(rkey(target["id"], region), res["ok"], res["stdout"], res["stderr"])
        if res["ok"]:
            print(f"[OK] {target['label']} / {region or '—'}")
        else:
            print(f"[ERROR] {target['label']} / {region or '—'}")
            if res.get("stderr"):
                print(res["stderr"][-1000:])
            exit_code = 1
    return exit_code


# =========================================================
# Helper functions
# =========================================================
def scan_output():
    files = []
    today = now_tw().date()
    if not OUTPUT_DIR.exists():
        return files

    for p in sorted(OUTPUT_DIR.rglob("*")):
        if p.is_file():
            s = p.stat()
            mdt = datetime.fromtimestamp(s.st_mtime, tz=TZ)
            files.append({
                "name": p.name,
                "path": p,
                "folder": str(p.parent.relative_to(OUTPUT_DIR)) if p.parent != OUTPUT_DIR else "根目錄",
                "size_kb": round(s.st_size / 1024, 1),
                "mtime": mdt,
                "mtime_str": mdt.strftime("%Y-%m-%d %H:%M"),
                "today": mdt.date() == today,
            })
    return sorted(files, key=lambda x: x["mtime"], reverse=True)


def read_last_lines(path, n=150):
    p = Path(path) if isinstance(path, str) else path
    if not p or not p.exists():
        return "(尚無 log 或檔案不存在)"
    try:
        lines = p.read_text(encoding="utf-8", errors="ignore").splitlines()
        return "\n".join(lines[-n:])
    except Exception as e:
        return f"(讀取失敗) {e}"


def highlight_log(text, kw=""):
    html = []
    for line in text.splitlines():
        esc = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if kw and kw.lower() in line.lower():
            esc = f'<mark style="background:#fef9c3;color:#713f12;border-radius:2px;padding:0 2px">{esc}</mark>'
        if any(k in line for k in ["Traceback", "Error", "ERROR", "❌", "PermissionError", "FAILED", "failed"]):
            html.append(f'<span class="le">{esc}</span>')
        elif any(k in line for k in ["✅", "SUCCESS", "success", "完成", "Done", "done", "[OK]"]):
            html.append(f'<span class="lo">{esc}</span>')
        elif any(k in line for k in ["WARNING", "Warning", "warn", "⚠", "[WARN]"]):
            html.append(f'<span class="lw">{esc}</span>')
        elif any(k in line for k in ["INFO", "info", "開始", "Start", "start", "[INFO]"]):
            html.append(f'<span class="li">{esc}</span>')
        else:
            html.append(f'<span class="ln">{esc}</span>')
    return "\n".join(html)


def is_today(entry: dict) -> bool:
    lr = entry.get("last_run", "")
    try:
        return datetime.strptime(lr[:10], "%Y-%m-%d").date() == now_tw().date()
    except Exception:
        return False


def badge(cls, label):
    return f'<span class="badge {cls}"><span class="dot"></span>{label}</span>'


def new_id(pfx="x"):
    return f"{pfx}{now_tw().strftime('%f')}"


# =========================================================
# CLI entry
# =========================================================
if IS_CLI_MODE:
    if "--list-jobs" in CLI_ARGS:
        list_jobs_cli()
        raise SystemExit(0)

    if "--run-job" in sys.argv:
        try:
            idx = sys.argv.index("--run-job")
            job_id = sys.argv[idx + 1]
        except Exception:
            print("[ERROR] Usage: python opapp.py --run-job <job_id>")
            raise SystemExit(1)
        raise SystemExit(run_one_job_cli(job_id))

    if "--run-scheduled" in CLI_ARGS:
        raise SystemExit(run_scheduled_jobs())


# =========================================================
# UI data load
# =========================================================
cfg = load_config()
runlog = load_runlog()
REGIONS = sorted(ACCOUNTS.keys()) if ACCOUNTS else []


# =========================================================
# UI
# =========================================================
now_str = now_tw().strftime("%Y/%m/%d  %H:%M")
rgn = st.session_state.region
chip = (
    f'<span class="rgn-chip">📍 {rgn}</span>'
    if rgn and rgn != REGION_ALL
    else '<span class="rgn-chip none">📍 未選地區</span>'
)
st.markdown(
    f'<div class="topbar"><div class="topbar-brand">📊 營運報表控制台</div>'
    f'<div class="topbar-sep"></div>{chip}'
    f'<div class="topbar-time">🕐 {now_str} 台北</div></div>',
    unsafe_allow_html=True,
)

PAGES = ["排程主控表", "手動執行", "Log 監控", "輸出報表", "腳本管理"]
ICONS = ["📋", "▶️", "📄", "📂", "⚙️"]
nav_cols = st.columns(len(PAGES) + 5)
for i, (pg, ic) in enumerate(zip(PAGES, ICONS)):
    wrap = "nav-wrap active" if st.session_state.page == pg else "nav-wrap"
    with nav_cols[i]:
        st.markdown(f'<div class="{wrap}">', unsafe_allow_html=True)
        if st.button(f"{ic} {pg}", key=f"nav_{pg}"):
            st.session_state.page = pg
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

page = st.session_state.page


def region_selector(allow_all=False):
    if not ACCOUNTS:
        st.error("⚠️ 找不到 Streamlit secrets 帳密，請到 App settings > Secrets 設定。")
        return None

    opts = ["（不指定地區）"]
    if allow_all:
        opts.append(REGION_ALL)
    opts += REGIONS

    cur = st.session_state.region
    idx = opts.index(cur) if cur in opts else 0

    cs, ci = st.columns([2, 5])
    with cs:
        sel = st.selectbox("📍 操作地區", opts, index=idx, key="rgn_widget")
    with ci:
        if sel not in ("（不指定地區）", REGION_ALL):
            if sel in ACCOUNTS:
                st.info(f"✅ **{sel}** 帳號：`{ACCOUNTS[sel].get('email', '—')}`")
            else:
                st.warning(f"⚠️ 找不到「{sel}」帳密")
        elif sel == REGION_ALL:
            st.info(f"🌐 將依序執行：{', '.join(REGIONS)}")
        else:
            st.caption("未指定地區時腳本不會收到帳密環境變數")

    nv = None if sel == "（不指定地區）" else sel
    if nv != st.session_state.region:
        st.session_state.region = nv
        st.rerun()
    return nv


def show_result(label, pairs):
    for rn, res in pairs:
        pre = f"**{rn}** · " if rn != "—" else ""
        if res["ok"]:
            st.success(f"✅ {pre}{label} 完成")
        else:
            st.error(f"❌ {pre}{label} 失敗")
        with st.expander(f"{'✅' if res['ok'] else '❌'} {pre}輸出詳情", expanded=not res["ok"]):
            st.caption(f"`{res.get('cmd', '')}`")
            if res.get("stdout"):
                st.text_area("stdout", res["stdout"], height=170, key=f"so{label}{rn}{id(res)}")
            if res.get("stderr"):
                st.text_area("stderr", res["stderr"], height=130, key=f"se{label}{rn}{id(res)}")


if page == "排程主控表":
    st.markdown('<div class="pg-title">排程主控表</div><div class="pg-sub">SCHEDULE DASHBOARD</div>', unsafe_allow_html=True)
    region = region_selector()

    all_jobs = cfg["daily"] + cfg["monthly"]
    enabled_jobs = [j for j in all_jobs if j.get("enabled", True)]
    ok_ct = sum(1 for j in enabled_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is True)
    fail_ct = sum(1 for j in enabled_jobs if (runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})).get("ok") is False)
    today_ct = sum(1 for j in enabled_jobs if is_today(runlog.get(rkey(j["id"], region)) or runlog.get(j["id"], {})))
    pend_ct = len(enabled_jobs) - ok_ct - fail_ct

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi blue">
        <div class="kpi-label">Total Tasks</div><div class="kpi-value">{len(all_jobs)}</div>
        <div class="kpi-sub">啟用 {len(enabled_jobs)} ／ 停用 {len(all_jobs)-len(enabled_jobs)}</div>
      </div>
      <div class="kpi green">
        <div class="kpi-label">今日已執行</div><div class="kpi-value">{today_ct}</div>
        <div class="kpi-sub">今日有 run log 紀錄</div>
      </div>
      <div class="kpi red">
        <div class="kpi-label">Failed</div><div class="kpi-value">{fail_ct}</div>
        <div class="kpi-sub">上次執行失敗</div>
      </div>
      <div class="kpi amber">
        <div class="kpi-label">Pending</div><div class="kpi-value">{pend_ct}</div>
        <div class="kpi-sub">尚未有執行紀錄</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    def render_section(grp_key, title, sec):
        jobs = cfg[grp_key]
        st.markdown(
            f'<div class="panel"><div class="panel-head"><div class="panel-tag">{title}</div>'
            f'<div class="panel-note">排程欄可直接點 ✏️ 編輯 · 勾選後可批次執行</div></div>',
            unsafe_allow_html=True,
        )

        hdr = st.columns([0.35, 0.35, 1.9, 1.9, 1.9, 1.8, 1.7, 0.85])
        for t, c in zip(["", "啟用", "任務名稱", "排程時間", "腳本", "狀態", "最後執行", "執行"], hdr):
            c.markdown(f'<div class="th">{t}</div>', unsafe_allow_html=True)

        checked = []
        for idx, job in enumerate(jobs):
            enabled = job.get("enabled", True)
            k = rkey(job["id"], region)
            entry = runlog.get(k) or runlog.get(job["id"], {})
            ok = entry.get("ok", None)
            last = entry.get("last_run", "—")
            today_ran = is_today(entry)
            exists = (BASE_DIR / job["script"]).exists()

            if ok is True and today_ran:
                st_b = badge("green", "今日成功")
            elif ok is True:
                st_b = badge("teal", "✓ 成功")
            elif ok is False and today_ran:
                st_b = badge("red", "今日失敗")
            elif ok is False:
                st_b = badge("red", "✗ 失敗")
            else:
                st_b = badge("gray", "待執行")

            e_icon = "🟢" if exists else "🔴"
            args_str = " ".join(job.get("args", []))
            script_t = f"{e_icon} `{job['script']}`" + (f" `{args_str}`" if args_str else "")
            edit_key = f"editing_{job['id']}"

            row = st.columns([0.35, 0.35, 1.9, 1.9, 1.9, 1.8, 1.7, 0.85])

            with row[0]:
                chk = st.checkbox("", key=f"sel_{sec}_{job['id']}", label_visibility="collapsed", disabled=not enabled)
                if chk and enabled:
                    checked.append(job)

            with row[1]:
                new_en = st.checkbox("", value=enabled, key=f"en_{sec}_{job['id']}", label_visibility="collapsed")
                if new_en != enabled:
                    cfg[grp_key][idx]["enabled"] = new_en
                    save_config(cfg)
                    st.rerun()

            with row[2]:
                name_md = f"**{job['label']}**" if enabled else f"<span style='color:var(--text-4)'>{job['label']}</span>"
                st.markdown(name_md, unsafe_allow_html=True)

            with row[3]:
                if st.session_state.get(edit_key, False):
                    nv = st.text_input(
                        "",
                        value=job.get("schedule", ""),
                        key=f"sv_{job['id']}",
                        placeholder="HH:MM / 每月15日18:15 / 月底18:30",
                        label_visibility="collapsed",
                    )
                    sc1, sc2 = st.columns(2)
                    with sc1:
                        if st.button("✓ 存", key=f"oks_{job['id']}"):
                            cfg[grp_key][idx]["schedule"] = nv
                            save_config(cfg)
                            st.session_state[edit_key] = False
                            st.rerun()
                    with sc2:
                        if st.button("✕ 取", key=f"cns_{job['id']}"):
                            st.session_state[edit_key] = False
                            st.rerun()
                else:
                    sv = job.get("schedule", "—")
                    nr = calc_next_run(sv)
                    nr_t = f'<div class="next-time">↻ {nr}</div>' if nr != "—" else ""
                    st.markdown(f'<div class="sched-lbl">⏰ {sv}</div>{nr_t}', unsafe_allow_html=True)
                    if st.button("✏️", key=f"eds_{job['id']}", help="編輯排程時間"):
                        st.session_state[edit_key] = True
                        st.rerun()

            with row[4]:
                st.markdown(script_t)

            with row[5]:
                st.markdown(st_b, unsafe_allow_html=True)

            with row[6]:
                if today_ran:
                    t_str = last[11:16] if len(last) >= 16 else ""
                    st.markdown(f'<span style="color:var(--green);font-size:12.5px;font-weight:700">✅ 今日 {t_str}</span>', unsafe_allow_html=True)
                else:
                    st.caption(last[:16] if last != "—" else "—")

            with row[7]:
                if enabled and st.button("▶", key=f"run_{sec}_{job['id']}", help=f"執行 {job['label']}"):
                    with st.spinner(f"執行 {job['label']} …"):
                        pairs = do_run_job(job, region)
                    show_result(job["label"], pairs)
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)
        return checked

    cd = render_section("daily", "📅 每日排程", "daily")
    cm = render_section("monthly", "🗓️ 月排程", "monthly")
    all_checked = cd + cm

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">⚙️ 批次控制</div></div>', unsafe_allow_html=True)
    cc1, cc2, cc3, cc4, _ = st.columns([1, 1.4, 1.5, 1.5, 2])
    with cc1:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()
    with cc2:
        if st.button("▶ 執行勾選", use_container_width=True, type="primary"):
            if not all_checked:
                st.warning("請先勾選至少一個任務")
            else:
                for job in all_checked:
                    with st.spinner(f"執行 {job['label']} …"):
                        pairs = do_run_job(job, region)
                    show_result(job["label"], pairs)
                st.rerun()
    with cc3:
        if st.button("▶ 全部每日報表", use_container_width=True):
            for job in [j for j in cfg["daily"] if j.get("enabled", True)]:
                with st.spinner(f"執行 {job['label']} …"):
                    pairs = do_run_job(job, region)
                show_result(job["label"], pairs)
            st.rerun()
    with cc4:
        if st.button("▶ 全部月報表", use_container_width=True):
            for job in [j for j in cfg["monthly"] if j.get("enabled", True)]:
                with st.spinner(f"執行 {job['label']} …"):
                    pairs = do_run_job(job, region)
                show_result(job["label"], pairs)
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
            with st.expander(f"❌ {job['label']}  ·  {entry.get('last_run', '—')}"):
                if entry.get("stderr"):
                    st.code(entry["stderr"], language="bash")
                if entry.get("stdout"):
                    st.code(entry["stdout"])
        st.markdown("</div>", unsafe_allow_html=True)

elif page == "手動執行":
    st.markdown('<div class="pg-title">手動執行</div><div class="pg-sub">MANUAL TRIGGER</div>', unsafe_allow_html=True)
    region = region_selector(allow_all=True)

    def render_manual(jobs, title):
        st.markdown(f'<div class="panel"><div class="panel-head"><div class="panel-tag">{title}</div><div class="panel-note">可輸入額外臨時參數</div></div>', unsafe_allow_html=True)
        for job in jobs:
            icon = "🟢" if (BASE_DIR / job["script"]).exists() else "🔴"
            en = job.get("enabled", True)
            c1, c2, c3, c4 = st.columns([1.8, 1.4, 2.8, 1.1])
            with c1:
                st.markdown(f"**{icon} {job['label']}**")
                st.caption(job["script"])
            with c2:
                sv = job.get("schedule", "—")
                nr = calc_next_run(sv)
                st.caption(f"排程：{sv}")
                if nr != "—":
                    st.caption(f"↻ {nr}")
            with c3:
                ea = st.text_input("額外參數（選填）", "", key=f"ea_{job['id']}", placeholder="如 202504 或留空", label_visibility="visible")
            with c4:
                st.markdown("<br>", unsafe_allow_html=True)
                lbl = "▶ 執行" if en else "⏸ 已停用"
                if st.button(lbl, key=f"man_{job['id']}", disabled=not en, use_container_width=True):
                    extra = ea.split() if ea.strip() else None
                    with st.spinner(f"執行 {job['label']} …"):
                        pairs = do_run_job(job, region, extra)
                    show_result(job["label"], pairs)
            st.divider()
        st.markdown("</div>", unsafe_allow_html=True)

    render_manual(cfg["daily"], "📅 每日報表")
    render_manual(cfg["monthly"], "🗓️ 月報表")

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">🚀 一鍵批次</div></div>', unsafe_allow_html=True)
    bc1, bc2, bc3 = st.columns(3)

    def batch(jobs, title):
        for job in [j for j in jobs if j.get("enabled", True)]:
            with st.spinner(f"執行 {job['label']} …"):
                pairs = do_run_job(job, region)
            show_result(job["label"], pairs)

    with bc1:
        if st.button("▶ 全部每日", use_container_width=True):
            batch(cfg["daily"], "每日")
    with bc2:
        if st.button("▶ 全部月報", use_container_width=True):
            batch(cfg["monthly"], "月報")
    with bc3:
        if st.button("▶ 全部任務", use_container_width=True):
            batch(cfg["daily"] + cfg["monthly"], "全部")
    st.markdown("</div>", unsafe_allow_html=True)

elif page == "Log 監控":
    st.markdown('<div class="pg-title">Log 監控</div><div class="pg-sub">LOG MONITOR</div>', unsafe_allow_html=True)

    builtin_logs = {"📋 run_log（JSON 執行紀錄）": RUNLOG_F}
    if LOG_DIR.exists():
        for lf in sorted(LOG_DIR.glob("*.log")):
            builtin_logs[f"📄 {lf.name}"] = lf
    for nm, ps in cfg.get("log_files", {}).items():
        builtin_logs[f"🔗 {nm}"] = Path(ps)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📄 Log 查看器</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns([2.5, 1, 2, 0.8])
    with c1:
        sel_log = st.selectbox("選擇 Log 檔", list(builtin_logs.keys()), key="log_sel")
    with c2:
        n_lines = st.selectbox("行數", [50, 100, 200, 500], index=1, key="log_n")
    with c3:
        kw = st.text_input("🔍 關鍵字篩選 / 高亮", placeholder="輸入後 Enter…", key="log_kw")
    with c4:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 刷新", use_container_width=True):
            st.rerun()

    lpath = builtin_logs[sel_log]

    if sel_log.startswith("📋"):
        st.caption(f"檔案：{lpath}")
        if lpath.exists():
            try:
                import pandas as pd
                data = json.loads(lpath.read_text(encoding="utf-8"))
                rows = [{
                    "任務 key": k,
                    "最後執行": v.get("last_run", "—"),
                    "狀態": "✅ 成功" if v.get("ok") else "❌ 失敗",
                    "stderr 預覽": (v.get("stderr") or "")[:100],
                } for k, v in sorted(data.items(), key=lambda x: x[1].get("last_run", ""), reverse=True)]
                df = pd.DataFrame(rows)
                if kw:
                    mask = df.apply(lambda row: row.astype(str).str.contains(kw, case=False).any(), axis=1)
                    df = df[mask]
                st.dataframe(df, use_container_width=True, hide_index=True)
                with st.expander("完整 JSON"):
                    st.json(data)
            except Exception as e:
                st.error(f"無法解析：{e}")
        else:
            st.info("尚無執行紀錄")
    else:
        mt = datetime.fromtimestamp(lpath.stat().st_mtime, tz=TZ).strftime("%Y-%m-%d %H:%M:%S") if lpath.exists() else "—"
        st.caption(f"檔案：{lpath}  ·  更新：{mt}")
        raw = read_last_lines(lpath, n_lines)
        if kw:
            lines = [l for l in raw.splitlines() if kw.lower() in l.lower()]
            raw = "\n".join(lines) if lines else f"（無符合「{kw}」的行）"
        st.markdown(f'<div class="logbox">{highlight_log(raw, kw)}</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">➕ 新增 Log 路徑</div></div>', unsafe_allow_html=True)
    with st.form("add_log"):
        l1, l2 = st.columns(2)
        with l1:
            alias = st.text_input("別名", placeholder="例：cron.log")
        with l2:
            fpath = st.text_input("完整路徑", placeholder="/path/to/file.log")
        if st.form_submit_button("➕ 加入"):
            if alias and fpath:
                cfg["log_files"][alias] = fpath
                save_config(cfg)
                st.success(f"已加入：{alias}")
                st.rerun()
            else:
                st.error("必填")
    extra = cfg.get("log_files", {})
    if extra:
        for nm in list(extra.keys()):
            r1, r2 = st.columns([5, 1])
            r1.caption(f"`{nm}` → {extra[nm]}")
            with r2:
                if st.button("移除", key=f"rm_{nm}"):
                    del cfg["log_files"][nm]
                    save_config(cfg)
                    st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

elif page == "輸出報表":
    st.markdown('<div class="pg-title">輸出報表</div><div class="pg-sub">OUTPUT FILES</div>', unsafe_allow_html=True)
    import pandas as pd

    files = scan_output()
    folders = sorted({f["folder"] for f in files})
    today_files = [f for f in files if f["today"]]
    total_kb = sum(f["size_kb"] for f in files)

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi blue"><div class="kpi-label">Total Files</div><div class="kpi-value">{len(files)}</div><div class="kpi-sub">output/ 下所有檔案</div></div>
      <div class="kpi green"><div class="kpi-label">今日產出</div><div class="kpi-value">{len(today_files)}</div><div class="kpi-sub">今日修改的檔案</div></div>
      <div class="kpi amber"><div class="kpi-label">Total Size</div><div class="kpi-value">{total_kb:.0f}</div><div class="kpi-sub">KB</div></div>
      <div class="kpi blue"><div class="kpi-label">Folders</div><div class="kpi-value">{len(folders)}</div><div class="kpi-sub">子目錄數</div></div>
    </div>
    """, unsafe_allow_html=True)

    if today_files:
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">🟢 今日產出</div></div>', unsafe_allow_html=True)
        st.dataframe(
            pd.DataFrame([{
                "檔名": f["name"],
                "資料夾": f["folder"],
                "時間": f["mtime_str"],
                "大小 KB": f["size_kb"],
            } for f in today_files[:30]]),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📁 全部檔案</div></div>', unsafe_allow_html=True)
    cf, ck, _, cr = st.columns([2, 3, 0.5, 1])
    with cf:
        sf = st.selectbox("資料夾", ["（全部）"] + folders, key="of")
    with ck:
        kw2 = st.text_input("🔍 搜尋檔名", placeholder="關鍵字…", key="ok")
    with cr:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 掃描", use_container_width=True):
            st.rerun()

    filtered = [
        f for f in files
        if (sf == "（全部）" or f["folder"] == sf)
        and (not kw2 or kw2.lower() in f["name"].lower())
    ]
    if not filtered:
        st.info("沒有符合條件的檔案")
    else:
        st.dataframe(
            pd.DataFrame([{
                "檔名": f["name"],
                "資料夾": f["folder"],
                "修改時間": f["mtime_str"],
                "大小 KB": f["size_kb"],
                "今日": "✅" if f["today"] else "—",
            } for f in filtered]),
            use_container_width=True,
            hide_index=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

elif page == "腳本管理":
    st.markdown('<div class="pg-title">腳本管理</div><div class="pg-sub">SCRIPT MANAGEMENT</div>', unsafe_allow_html=True)
    tab_d, tab_m, tab_e, tab_r = st.tabs(["📅 每日腳本", "🗓️ 月腳本", "📝 線上編輯", "🔧 重置"])

    def job_editor(grp_key, tab):
        with tab:
            jobs = cfg[grp_key]
            st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">現有腳本</div></div>', unsafe_allow_html=True)
            for idx, job in enumerate(jobs):
                ex = (BASE_DIR / job["script"]).exists()
                icon = "🟢" if ex else "🔴"
                en_lbl = "✅ 啟用" if job.get("enabled", True) else "⏸ 停用"
                with st.expander(f"{icon} {job['label']}  ·  {job.get('schedule', '—')}  ·  {en_lbl}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        nl = st.text_input("顯示名稱", job["label"], key=f"lbl_{job['id']}")
                        ns = st.text_input("腳本檔名", job["script"], key=f"scr_{job['id']}")
                    with c2:
                        na = st.text_input("固定參數", " ".join(job.get("args", [])), key=f"arg_{job['id']}")
                        nt = st.text_input("排程時間（HH:MM / 每月15日18:15 / 月底18:30）", job.get("schedule", ""), key=f"sch_{job['id']}")
                    c3, c4 = st.columns(2)
                    with c3:
                        nm = st.checkbox("需跑全部地區", value=job.get("all_regions", False), key=f"mul_{job['id']}")
                    with c4:
                        ne = st.checkbox("啟用此任務", value=job.get("enabled", True), key=f"en2_{job['id']}")
                    nr = calc_next_run(nt)
                    if nr != "—":
                        st.markdown(f'<div class="next-run">⏭️ 下次執行預估：<strong>{nr}</strong></div>', unsafe_allow_html=True)
                    a1, a2, a3 = st.columns(3)
                    with a1:
                        if st.button("💾 儲存", key=f"sv_{job['id']}"):
                            cfg[grp_key][idx].update({
                                "label": nl,
                                "script": ns,
                                "args": na.split() if na.strip() else [],
                                "schedule": nt,
                                "all_regions": nm,
                                "enabled": ne,
                            })
                            save_config(cfg)
                            st.success("✅ 已儲存")
                            st.rerun()
                    with a2:
                        if st.button("▶ 測試", key=f"ts_{job['id']}"):
                            with st.spinner("…"):
                                pairs = do_run_job(job, st.session_state.region)
                            show_result(job["label"], pairs)
                    with a3:
                        if st.button("🗑️ 刪除", key=f"dl_{job['id']}"):
                            cfg[grp_key].pop(idx)
                            save_config(cfg)
                            st.warning(f"已刪除 {job['label']}")
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">➕ 新增腳本</div></div>', unsafe_allow_html=True)
            with st.form(f"add_{grp_key}"):
                c1, c2 = st.columns(2)
                with c1:
                    fl = st.text_input("顯示名稱", placeholder="例：新報表")
                    fs = st.text_input("腳本檔名", placeholder="例：new_report.py")
                with c2:
                    fa = st.text_input("固定參數", placeholder="留空或 1")
                    ft = st.text_input("排程時間", placeholder="08:00 / 每月15日18:15 / 月底18:30")
                fc1, fc2 = st.columns(2)
                with fc1:
                    fm = st.checkbox("需跑全部地區")
                with fc2:
                    fen = st.checkbox("啟用", value=True)
                if st.form_submit_button("➕ 新增"):
                    if not fl or not fs:
                        st.error("名稱與腳本必填")
                    else:
                        cfg[grp_key].append({
                            "id": new_id(grp_key[0]),
                            "label": fl,
                            "script": fs,
                            "args": fa.split() if fa.strip() else [],
                            "schedule": ft or "—",
                            "all_regions": fm,
                            "enabled": fen,
                        })
                        save_config(cfg)
                        st.success(f"✅ 已新增：{fl}")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    job_editor("daily", tab_d)
    job_editor("monthly", tab_m)

    with tab_e:
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">📝 直接編輯腳本</div></div>', unsafe_allow_html=True)
        scripts = sorted({j["script"] for grp in ["daily", "monthly"] for j in cfg[grp]})
        sel = st.selectbox("選擇腳本", ["（請選擇）"] + scripts, key="edit_sel")
        if sel and sel != "（請選擇）":
            sp = BASE_DIR / sel
            st.caption(f"{'✅ 存在' if sp.exists() else '🔴 不存在（儲存後建立）'}  ·  {sp}")
            code = sp.read_text(encoding="utf-8") if sp.exists() else f"# {sel}\n"
            edited = st.text_area("", code, height=480, key="ec", label_visibility="collapsed")
            e1, e2 = st.columns(2)
            with e1:
                if st.button("💾 寫入儲存", type="primary", use_container_width=True):
                    sp.write_text(edited, encoding="utf-8")
                    st.success(f"✅ 已寫入 {sel}")
            with e2:
                if st.button("🔄 重新讀取", use_container_width=True):
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_r:
        st.markdown('<div class="panel"><div class="panel-head"><div class="panel-tag">⚠️ 重置設定</div></div>', unsafe_allow_html=True)
        st.warning("重置後所有排程設定（名稱、時間、腳本）還原為預設值，run_log 不受影響。")
        if st.checkbox("✅ 確認重置"):
            if st.button("🔄 執行重置"):
                save_config(_deepcopy_config(DEFAULT_CONFIG))
                st.warning("已重置，請重新整理")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.caption(
    f"營運報表控制台 v6.1  ·  {now_tw().strftime('%Y-%m-%d %H:%M:%S')} 台北  ·  "
    f"可用地區：{', '.join(REGIONS) if REGIONS else '尚未設定'}"
)
