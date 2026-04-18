import json
import plistlib
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, Dict, Callable, List

import pandas as pd
import streamlit as st

from paths import (
    PATH_CLEANER_DATA,
    PATH_CLEANER_SCHEDULE,
    PATH_HR,
    PATH_ORDER,
    PATH_REPORT,
    PATH_SCHEDULE,
    PATH_VIP,
)
from performance_report import (
    generate_sales_report,
    load_execution_log_for_current_month,
    delete_execution_log_rows,
    LATEST_DIR,
)

# ── Constants ─────────────────────────────────────────────────────────────────

TZ_TAIPEI = timezone(timedelta(hours=8))
PROJECT_DIR = Path(__file__).resolve().parent
LOCAL_BASE_DIR = Path("/Users/jenny/lemon")
BASE_DIR = LOCAL_BASE_DIR if LOCAL_BASE_DIR.exists() else PROJECT_DIR
IS_LOCAL = LOCAL_BASE_DIR.exists()

LOG_FILE = BASE_DIR / "cron.log"
LAUNCH_AGENTS_DIR = Path.home() / "Library/LaunchAgents"
CONFIG_FILE = BASE_DIR / "dashboard_config.json"

PYTHON_CMD = sys.executable or "python3"

OUTPUT_DIRS = {
    "排班統計表": Path(PATH_SCHEDULE),
    "專員班表": Path(PATH_CLEANER_SCHEDULE),
    "專員個資": Path(PATH_CLEANER_DATA),
    "訂單資料": Path(PATH_ORDER),
    "業績報表": Path(PATH_REPORT),
    "預收": Path(PATH_VIP),
    "儲值金結算": Path(PATH_VIP),
    "儲值金預收": Path(PATH_VIP),
    "上下半月訂單": Path(PATH_HR),
    "已退款": Path(PATH_HR),
}

MAIN_REPORT_TASKS = [
    {
        "name": "排班統計表",
        "task_key": "schedule_report",
        "script": "schedule_report.py",
        "label": "com.jenny.daily01",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01.plist",
        "default_hour": "01",
        "default_minute": "10",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" schedule_report.py',
    },
    {
        "name": "專員班表",
        "task_key": "staff_schedule",
        "script": "staff_schedule.py",
        "label": "com.jenny.daily01b",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01b.plist",
        "default_hour": "01",
        "default_minute": "20",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_schedule.py',
    },
    {
        "name": "專員個資",
        "task_key": "staff_info",
        "script": "staff_info.py",
        "label": "com.jenny.daily02b",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02b.plist",
        "default_hour": "01",
        "default_minute": "30",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_info.py',
    },
    {
        "name": "訂單資料",
        "task_key": "orders_report",
        "script": "orders_report.py",
        "label": "com.jenny.daily02",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02.plist",
        "default_hour": "01",
        "default_minute": "40",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" orders_report.py',
    },
    {
        "name": "業績報表",
        "task_key": "performance_report",
        "script": "performance_report.py",
        "label": "com.jenny.sales08",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.sales08.plist",
        "default_hour": "08",
        "default_minute": "00",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" performance_report.py dashboard false',
    },
]

MANUAL_TASKS = [
    {"name": "排班統計表", "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" schedule_report.py'},
    {"name": "專員班表",   "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_schedule.py'},
    {"name": "專員個資",   "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_info.py'},
    {"name": "訂單資料",   "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" orders_report.py'},
    {"name": "業績報表",   "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" performance_report.py dashboard false'},
]

NAV_PAGES = ["主控表", "業績報表", "上下半月訂單", "手動執行", "Log 監控", "輸出檔案", "程式管理", "排程設定"]
NAV_ICONS = ["📋", "💹", "🧾", "▶️", "📄", "📂", "⚙️", "⏰"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def now_taipei():
    return datetime.now(TZ_TAIPEI)


def run_shell(cmd: str):
    p = subprocess.run(cmd, shell=True, text=True, capture_output=True, executable="/bin/bash")
    return p.returncode, p.stdout, p.stderr


def read_last_lines(path: Path, n: int = 200) -> str:
    if not path.exists():
        return "(尚無 log)"
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return "".join(f.readlines()[-n:])
    except Exception as e:
        return f"(讀取失敗) {e}"


def file_mtime(path: Optional[Path]) -> str:
    if not path or not path.exists():
        return "-"
    return datetime.fromtimestamp(path.stat().st_mtime, tz=TZ_TAIPEI).strftime("%m/%d %H:%M")


def file_size_str(path: Optional[Path]) -> str:
    if not path or not path.exists():
        return "-"
    s = path.stat().st_size
    if s < 1024:        return f"{s} B"
    if s < 1024 * 1024: return f"{s/1024:.1f} KB"
    return f"{s/1024/1024:.1f} MB"


def get_launchd_status() -> dict:
    code, out, _ = run_shell("launchctl list | grep com.jenny")
    status_map = {}
    if code != 0 and not out.strip():
        return status_map
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 3:
            status_map[parts[2]] = {"pid": parts[0], "last_exit": parts[1]}
    return status_map


def launchd_badge(label: str, status_map: dict) -> tuple[str, str]:
    info = status_map.get(label)
    if not info:
        return ("未載入", "gray")
    pid, ex = info["pid"], info["last_exit"]
    if pid != "-":
        return (f"執行中 {pid}", "yellow")
    return ("正常", "green") if ex == "0" else (f"異常 {ex}", "red")


def load_plist_schedule(plist_path: Path, default_hour="", default_minute="") -> dict:
    fallback = {"exists": False, "hour": default_hour, "minute": default_minute, "day": "", "source": "default"}
    if not plist_path.exists():
        return fallback
    try:
        with open(plist_path, "rb") as f:
            data = plistlib.load(f)
        iv = data.get("StartCalendarInterval", {})
        if isinstance(iv, list):
            iv = iv[0] if iv else {}
        def _z(v, fb):
            return str(v).zfill(2) if str(v) != "" else fb
        return {
            "exists": True,
            "hour": _z(iv.get("Hour", default_hour), default_hour),
            "minute": _z(iv.get("Minute", default_minute), default_minute),
            "day": str(iv.get("Day", "")),
            "source": "plist",
        }
    except Exception:
        return fallback


def save_plist_schedule(plist_path, hour, minute, day=""):
    with open(plist_path, "rb") as f:
        data = plistlib.load(f)
    iv = {}
    if day.strip():
        iv["Day"] = int(day)
    iv["Hour"] = int(hour)
    iv["Minute"] = int(minute)
    data["StartCalendarInterval"] = iv
    with open(plist_path, "wb") as f:
        plistlib.dump(data, f)
    run_shell(f'launchctl bootout gui/$(id -u) "{plist_path}" 2>/dev/null')
    return run_shell(f'launchctl bootstrap gui/$(id -u) "{plist_path}"')


def calc_next_run(day_str, hour_str, minute_str) -> str:
    try:
        hour, minute = int(hour_str), int(minute_str)
        now = now_taipei().replace(tzinfo=None)
        if str(day_str).strip():
            day = int(day_str)
            c = now.replace(day=day, hour=hour, minute=minute, second=0, microsecond=0)
            if c <= now:
                c = c.replace(month=now.month + 1) if now.month < 12 else c.replace(year=now.year + 1, month=1)
            return c.strftime("%Y-%m-%d %H:%M")
        c = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if c <= now:
            c += timedelta(days=1)
        return c.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return "—"


def find_latest_files(base_dir: Path, limit=10) -> List[Path]:
    if not base_dir.exists():
        return []
    files = []
    try:
        for p in base_dir.rglob("*"):
            if p.is_file() and not p.name.startswith((".", "~$")):
                files.append(p)
    except Exception:
        return []
    return sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)[:limit]


def load_sales_latest_payload() -> dict:
    latest_dir = Path(LATEST_DIR)
    payload = {"df4": pd.DataFrame(), "daily_df": pd.DataFrame(), "meta": {}, "email_html": ""}
    for key, fname, enc in [
        ("df4", "df4.csv", "utf-8-sig"),
        ("daily_df", "daily_df.csv", "utf-8-sig"),
    ]:
        p = latest_dir / fname
        if p.exists():
            try:
                payload[key] = pd.read_csv(p, encoding=enc)
            except Exception as e:
                payload[f"{key}_error"] = str(e)
    meta_p = latest_dir / "meta.json"
    if meta_p.exists():
        try:
            payload["meta"] = json.loads(meta_p.read_text(encoding="utf-8"))
        except Exception:
            pass
    html_p = latest_dir / "email_preview.html"
    if html_p.exists():
        payload["email_html"] = html_p.read_text(encoding="utf-8")
    return payload


def highlight_log(text: str) -> str:
    lines = []
    for line in text.splitlines():
        e = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if any(k in line for k in ["Traceback", "Error", "ERROR", "❌", "PermissionError", "FAILED"]):
            lines.append(f'<span class="log-err">{e}</span>')
        elif any(k in line for k in ["✅", "SUCCESS", "success", "完成", "Done"]):
            lines.append(f'<span class="log-ok">{e}</span>')
        elif any(k in line for k in ["WARNING", "warn", "⚠"]):
            lines.append(f'<span class="log-warn">{e}</span>')
        elif any(k in line for k in ["INFO", "開始", "Start"]):
            lines.append(f'<span class="log-info">{e}</span>')
        else:
            lines.append(f'<span class="log-normal">{e}</span>')
    return "\n".join(lines)


# ── HTML table helper (guarantees right-align) ────────────────────────────────

def _html_table(df: pd.DataFrame, right_cols: set, pct_cols: set, int_cols: set) -> str:
    """Render a DataFrame as a styled HTML table with proper column alignment."""

    def fmt_cell(val, col):
        if pd.isna(val) or val == "":
            return "—"
        if col in pct_cols:
            try:
                return f"{float(val):.2%}"
            except Exception:
                return str(val)
        if col in int_cols:
            try:
                return f"{int(float(val)):,}"
            except Exception:
                return str(val)
        return str(val)

    th_style = (
        "padding:10px 14px;"
        "font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;"
        "color:#64748b;border-bottom:2px solid #e2e8f0;white-space:nowrap;"
        "background:#fafafa;"
    )
    td_base = "padding:9px 14px;font-size:13px;color:#1e293b;border-bottom:1px solid #f1f5f9;white-space:nowrap;"
    td_right = td_base + "text-align:right;font-variant-numeric:tabular-nums;font-family:'DM Mono',monospace;"
    td_left  = td_base + "text-align:left;"

    rows_html = []

    # header
    ths = "".join(
        f'<th style="{th_style}text-align:{"right" if c in right_cols else "left"}">{c}</th>'
        for c in df.columns
    )
    rows_html.append(f"<thead><tr>{ths}</tr></thead>")

    # body
    body_rows = []
    for _, row in df.iterrows():
        cells = "".join(
            f'<td style="{td_right if c in right_cols else td_left}">{fmt_cell(row[c], c)}</td>'
            for c in df.columns
        )
        body_rows.append(f"<tr>{cells}</tr>")
    rows_html.append(f"<tbody>{''.join(body_rows)}</tbody>")

    table_style = (
        "width:100%;border-collapse:collapse;"
        "background:#fff;border-radius:12px;overflow:hidden;"
    )
    return f'<div style="overflow-x:auto;border:1.5px solid #e2e8f0;border-radius:12px;">' \
           f'<table style="{table_style}">{"".join(rows_html)}</table></div>'


# ── Global CSS ────────────────────────────────────────────────────────────────

GLOBAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
    background: #f0f2f6 !important;
    font-family: 'DM Sans','PingFang TC','Noto Sans TC',sans-serif !important;
    color: #1e293b !important;
}
[data-testid="stHeader"], [data-testid="stSidebar"] { display: none !important; }
.block-container { padding: 0 2.6rem 4rem !important; max-width: 1480px !important; }

/* ── Top nav ── */
.topnav-wrapper {
    background: #fff; margin: 0 -2.6rem; padding: 0 28px;
    border-bottom: 1.5px solid #e2e8f0; position: sticky; top: 0; z-index: 999;
    box-shadow: 0 2px 12px rgba(15,23,42,.07);
}
.topnav-top { display:flex; align-items:center; height:58px; }
.topnav-brand { display:flex; align-items:center; gap:10px; flex-shrink:0; }
.topnav-logo  { font-size:22px; line-height:1; }
.topnav-name  { font-size:15px; font-weight:700; color:#0f172a; letter-spacing:-.01em; white-space:nowrap; }
.topnav-divider { width:1.5px; height:24px; background:#e2e8f0; margin:0 22px; flex-shrink:0; }
.topnav-time  { font-size:12px; color:#64748b; font-weight:500; margin-left:auto; }
.topnav-py    { font-size:11px; color:#94a3b8; margin-left:12px; font-family:'DM Mono',monospace; }

/* ── Nav buttons ── */
.nav-item-wrap div[data-testid="stButton"] > button {
    height:44px !important; padding:0 16px !important; border-radius:0 !important;
    border:none !important; border-bottom:2.5px solid transparent !important;
    background:transparent !important; color:#64748b !important;
    font-weight:600 !important; font-size:12.5px !important; box-shadow:none !important;
}
.nav-item-wrap div[data-testid="stButton"] > button:hover { color:#334155 !important; background:#f8fafc !important; }
.nav-item-wrap.active div[data-testid="stButton"] > button { color:#2563eb !important; border-bottom:2.5px solid #2563eb !important; }

/* ── Page header ── */
.page-header {
    padding:26px 0 18px; border-bottom:1.5px solid #e2e8f0;
    margin-bottom:26px; display:flex; align-items:flex-end; gap:14px;
}
.page-title    { font-size:23px; font-weight:700; color:#0f172a; line-height:1; letter-spacing:-.02em; }
.page-subtitle { font-size:10.5px; font-weight:700; letter-spacing:.15em; text-transform:uppercase; color:#94a3b8; padding-bottom:2px; }

/* ── KPI cards ── */
.kpi-row { display:flex; gap:14px; margin-bottom:26px; }
.kpi-card {
    flex:1; background:#fff; border:1.5px solid #e2e8f0; border-radius:14px;
    padding:18px 22px 16px; position:relative; overflow:hidden;
    box-shadow:0 1px 4px rgba(15,23,42,.05),0 4px 14px rgba(15,23,42,.04);
}
.kpi-card::before { content:''; position:absolute; top:0; left:0; right:0; height:3.5px; border-radius:14px 14px 0 0; }
.kpi-card.blue::before  { background:linear-gradient(90deg,#2563eb,#60a5fa); }
.kpi-card.green::before { background:linear-gradient(90deg,#059669,#34d399); }
.kpi-card.amber::before { background:linear-gradient(90deg,#b45309,#fbbf24); }
.kpi-card.red::before   { background:linear-gradient(90deg,#dc2626,#f87171); }
.kpi-label { font-size:10px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:#64748b; margin-bottom:8px; }
.kpi-value { font-size:36px; font-weight:700; color:#0f172a; line-height:1; letter-spacing:-.03em; font-variant-numeric:tabular-nums; }
.kpi-sub   { font-size:12px; color:#64748b; font-weight:500; margin-top:6px; }

/* ── Section card ── */
.section-card {
    background:#fff; border:1.5px solid #e2e8f0; border-radius:14px;
    padding:22px 24px 20px; margin-bottom:18px;
    box-shadow:0 1px 3px rgba(15,23,42,.04),0 4px 14px rgba(15,23,42,.04);
}
.section-title {
    font-size:13px; font-weight:700; letter-spacing:.08em; text-transform:uppercase;
    color:#2563eb; margin-bottom:18px; padding-bottom:12px; border-bottom:1px solid #f1f5f9;
    display:flex; align-items:center; gap:8px;
}

/* ── Task row (主控表) ── */
.task-header {
    display:grid;
    grid-template-columns:1.6fr 1fr 1fr 1fr 1.1fr 0.6fr;
    gap:0; padding:0 4px 10px;
    border-bottom:1.5px solid #e2e8f0; margin-bottom:6px;
}
.task-header span {
    font-size:10px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:#94a3b8;
}

/* ── Status badges ── */
.badge {
    display:inline-flex; align-items:center; gap:5px;
    font-size:11.5px; font-weight:600; padding:3px 10px; border-radius:20px; white-space:nowrap;
}
.badge::before { content:''; width:6px; height:6px; border-radius:50%; flex-shrink:0; }
.b-green  { color:#065f46; background:#d1fae5; } .b-green::before  { background:#059669; }
.b-yellow { color:#78350f; background:#fef3c7; } .b-yellow::before { background:#d97706; }
.b-red    { color:#991b1b; background:#fee2e2; } .b-red::before    { background:#dc2626; }
.b-gray   { color:#475569; background:#f1f5f9; } .b-gray::before   { background:#94a3b8; }

/* ── Run button (small) ── */
.run-btn-wrap div[data-testid="stButton"] > button {
    background:#1e293b !important; color:#f1f5f9 !important;
    border:none !important; border-radius:7px !important;
    font-weight:600 !important; font-size:12px !important;
    padding:4px 12px !important; height:30px !important; min-height:30px !important; box-shadow:none !important;
}
.run-btn-wrap div[data-testid="stButton"] > button:hover { background:#0f172a !important; }

/* ── Save btn (tiny) ── */
.save-btn-wrap div[data-testid="stButton"] > button {
    background:#e0f2fe !important; color:#0369a1 !important;
    border:1px solid #bae6fd !important; border-radius:6px !important;
    font-weight:700 !important; font-size:12px !important;
    padding:3px 10px !important; height:28px !important; min-height:28px !important; box-shadow:none !important;
}

/* ── Exec result panel ── */
.exec-panel {
    background:#fff; border:1.5px solid #e2e8f0; border-radius:12px;
    padding:16px 20px; margin-top:12px;
    box-shadow:0 1px 3px rgba(15,23,42,.04);
}
.exec-panel-title { font-size:12.5px; font-weight:700; color:#0f172a; margin-bottom:10px; display:flex; align-items:center; gap:8px; }
.exec-label { font-size:9.5px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:#94a3b8; margin:10px 0 5px; }

/* ── Log box ── */
.log-box {
    background:#0d1117; border:1px solid #1e2d3d; border-radius:10px;
    padding:14px 18px; font-family:'DM Mono','Menlo',monospace; font-size:12.5px;
    line-height:1.75; white-space:pre-wrap; word-break:break-all; max-height:420px; overflow:auto;
}
.log-err    { color:#f87171; display:block; }
.log-ok     { color:#4ade80; display:block; }
.log-warn   { color:#fbbf24; display:block; }
.log-info   { color:#60a5fa; display:block; }
.log-normal { color:#94a3b8; display:block; }
.log-meta   { font-size:11.5px; color:#64748b; font-weight:500; margin-bottom:8px; }

/* ── Streamlit overrides ── */
div[data-testid="stButton"] > button {
    background:#1e293b !important; color:#f8fafc !important; border:none !important;
    border-radius:8px !important; font-weight:600 !important; font-size:13px !important;
    padding:8px 18px !important; box-shadow:0 1px 3px rgba(15,23,42,.12) !important;
}
div[data-testid="stButton"] > button:hover { background:#0f172a !important; }
div[data-testid="stSelectbox"] > div > div,
div[data-testid="stTextInput"] > div > div > input {
    background:#fff !important; border:1.5px solid #cbd5e1 !important;
    border-radius:8px !important; color:#1e293b !important; font-size:13.5px !important;
}
div[data-testid="stSelectbox"] label,
div[data-testid="stTextInput"] label,
div[data-testid="stTextArea"] label { color:#374151 !important; font-size:13px !important; font-weight:600 !important; }
div[data-testid="stTextArea"] textarea {
    border:1.5px solid #cbd5e1 !important; border-radius:8px !important;
    font-family:'DM Mono',monospace !important; font-size:12.5px !important;
    color:#1e293b !important; background:#fafafa !important;
}
div[data-testid="stMetric"] {
    background:#fff; border-radius:12px; padding:16px 18px;
    border:1.5px solid #e2e8f0; box-shadow:0 1px 3px rgba(15,23,42,.04);
}
div[data-testid="stMetric"] label { color:#475569 !important; font-size:12px !important; font-weight:600 !important; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { color:#0f172a !important; font-size:28px !important; font-weight:700 !important; }
div[data-testid="stDataFrame"] { border-radius:10px !important; overflow:hidden !important; border:1.5px solid #e2e8f0 !important; }
div[data-testid="stAlert"] { border-radius:10px !important; font-size:13px !important; font-weight:500 !important; }
.stCaption, div[data-testid="stCaption"] { color:#64748b !important; font-size:12px !important; font-weight:500 !important; }

/* ── Divider ── */
.divider { height:1.5px; background:#e2e8f0; margin:20px 0; }

/* ── Next-run pill ── */
.next-run { display:inline-flex; align-items:center; gap:5px; background:#f8fafc; border:1.5px solid #e2e8f0; border-radius:8px; padding:4px 10px; font-size:11.5px; font-weight:600; color:#475569; }

/* ── Footer ── */
.footer-cap { text-align:center; font-size:11px; color:#94a3b8; font-weight:500; padding-top:28px; border-top:1.5px solid #e2e8f0; margin-top:32px; }

/* ── Data empty state ── */
.empty-state { text-align:center; padding:32px 20px; color:#94a3b8; font-size:13px; font-weight:500; background:#f8fafc; border-radius:10px; border:1.5px dashed #e2e8f0; }
.empty-state .icon { font-size:28px; display:block; margin-bottom:8px; }
</style>
"""


# ── Streamlit setup ───────────────────────────────────────────────────────────

st.set_page_config(page_title="Jenny 排程控制台", page_icon="🍋", layout="wide")

if "page" not in st.session_state:
    st.session_state.page = "主控表"
if "task_results" not in st.session_state:
    st.session_state.task_results = {}

st.markdown(GLOBAL_CSS, unsafe_allow_html=True)

# Top bar
now_str = now_taipei().strftime("%Y/%m/%d  %H:%M")
st.markdown(
    f"""<div class="topnav-wrapper">
      <div class="topnav-top">
        <div class="topnav-brand"><span class="topnav-logo">🍋</span><span class="topnav-name">Jenny 排程控制台</span></div>
        <div class="topnav-divider"></div>
        <div class="topnav-time">🕐 {now_str}<span class="topnav-py">{PYTHON_CMD}</span></div>
      </div>
    </div>""",
    unsafe_allow_html=True,
)

# Navigation
nav_cols = st.columns(len(NAV_PAGES))
for i, (label, icon) in enumerate(zip(NAV_PAGES, NAV_ICONS)):
    active = st.session_state.page == label
    with nav_cols[i]:
        st.markdown(f'<div class="nav-item-wrap {"active" if active else ""}">', unsafe_allow_html=True)
        if st.button(f"{icon} {label}", key=f"nav_{label}"):
            st.session_state.page = label
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

page = st.session_state.page

CLS = {"green": "b-green", "yellow": "b-yellow", "red": "b-red", "gray": "b-gray"}
def badge(label, cls): return f'<span class="badge {CLS.get(cls,"b-gray")}">{label}</span>'


# ── 主控表 ────────────────────────────────────────────────────────────────────

if page == "主控表":
    st.markdown(
        '<div class="page-header"><div class="page-title">排程主控表</div><div class="page-subtitle">Schedule Dashboard</div></div>',
        unsafe_allow_html=True,
    )

    status_map = get_launchd_status()
    count_ok = sum(1 for t in MAIN_REPORT_TASKS if launchd_badge(t["label"], status_map)[1] == "green")
    count_run = sum(1 for t in MAIN_REPORT_TASKS if launchd_badge(t["label"], status_map)[1] == "yellow")
    count_err = len(MAIN_REPORT_TASKS) - count_ok - count_run
    ran_today = sum(1 for k, v in st.session_state.task_results.items() if v)

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi-card blue"><div class="kpi-label">Total Tasks</div><div class="kpi-value">{len(MAIN_REPORT_TASKS)}</div><div class="kpi-sub">已設定排程</div></div>
      <div class="kpi-card green"><div class="kpi-label">Normal</div><div class="kpi-value">{count_ok}</div><div class="kpi-sub">上次退出正常</div></div>
      <div class="kpi-card amber"><div class="kpi-label">Running</div><div class="kpi-value">{count_run}</div><div class="kpi-sub">目前執行中</div></div>
      <div class="kpi-card {"red" if count_err > 0 else "blue"}"><div class="kpi-label">今日已執行</div><div class="kpi-value">{ran_today}</div><div class="kpi-sub">本次 session</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 報表任務</div>', unsafe_allow_html=True)

    # column header
    st.markdown("""
    <div style="display:grid;grid-template-columns:1.5fr 0.8fr 1.4fr 1fr 1fr 1.1fr 0.55fr;
                gap:0;padding:0 4px 10px;border-bottom:1.5px solid #e2e8f0;margin-bottom:2px;">
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">任務 / 腳本</span>
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">排程狀態</span>
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">修改排程時間</span>
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">執行狀態</span>
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">上次結果</span>
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">下次執行</span>
      <span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;text-align:center;">執行</span>
    </div>
    """, unsafe_allow_html=True)

    for task in MAIN_REPORT_TASKS:
        sched = load_plist_schedule(task["plist"], task["default_hour"], task["default_minute"])
        ld_text, ld_cls = launchd_badge(task["label"], status_map)
        result_data = st.session_state.task_results.get(task["task_key"])

        # result badge
        if result_data is None:
            res_badge = badge("尚未執行", "gray")
        elif result_data["code"] == 0:
            res_badge = badge("✓ 成功", "green")
        else:
            res_badge = badge("✗ 失敗", "red")

        c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 0.8, 1.4, 1, 1, 1.1, 0.55])

        with c1:
            st.markdown(
                f"<span style='font-weight:700;color:#0f172a;font-size:13px'>{task['name']}</span><br>"
                f"<span style='font-size:11px;color:#94a3b8;font-family:monospace'>{task['script']}</span>",
                unsafe_allow_html=True,
            )

        with c2:
            st.markdown(badge(ld_text, ld_cls), unsafe_allow_html=True)

        with c3:
            ci1, ci2, ci3 = st.columns([1, 1, 0.7])
            with ci1:
                h_val = st.text_input("時", value=sched["hour"], key=f'h_{task["task_key"]}', label_visibility="collapsed", placeholder="HH")
            with ci2:
                m_val = st.text_input("分", value=sched["minute"], key=f'm_{task["task_key"]}', label_visibility="collapsed", placeholder="MM")
            with ci3:
                st.markdown('<div class="save-btn-wrap">', unsafe_allow_html=True)
                if st.button("💾", key=f'save_{task["task_key"]}', use_container_width=True):
                    if not IS_LOCAL:
                        st.warning("雲端環境無法修改本機 plist")
                    elif not task["plist"].exists():
                        st.error(f"找不到 plist：{task['plist'].name}")
                    elif not h_val.isdigit() or not m_val.isdigit():
                        st.error("時間必須是數字")
                    else:
                        code, out, err = save_plist_schedule(task["plist"], h_val, m_val, sched["day"])
                        if code == 0:
                            st.success("✓ 已更新")
                            st.rerun()
                        else:
                            st.error(err or "更新失敗")
                st.markdown("</div>", unsafe_allow_html=True)

        with c4:
            note = f'<span style="font-size:10.5px;color:#94a3b8">({sched["source"]})</span>' if sched["source"] == "default" else ""
            t_h = sched.get("hour", "?")
            t_m = sched.get("minute", "?")
            st.markdown(
                f'<span style="font-family:monospace;font-size:13px;font-weight:600;color:#1e293b">{t_h}:{t_m}</span> {note}',
                unsafe_allow_html=True,
            )

        with c5:
            st.markdown(res_badge, unsafe_allow_html=True)
            if result_data:
                st.markdown(
                    f'<span style="font-size:10.5px;color:#94a3b8">{result_data.get("ran_at","")[-8:]}</span>',
                    unsafe_allow_html=True,
                )

        with c6:
            next_t = calc_next_run(sched["day"], sched["hour"], sched["minute"])
            st.markdown(f'<span class="next-run">🕐 {next_t}</span>', unsafe_allow_html=True)

        with c7:
            st.markdown('<div class="run-btn-wrap">', unsafe_allow_html=True)
            if st.button("▶", key=f'run_{task["task_key"]}', help=f'執行 {task["name"]}'):
                with st.spinner(f"執行中：{task['name']}…"):
                    rc, out, err = run_shell(task["cmd"])
                st.session_state.task_results[task["task_key"]] = {
                    "name": task["name"],
                    "code": rc,
                    "stdout": out,
                    "stderr": err,
                    "ran_at": now_taipei().strftime("%Y-%m-%d %H:%M:%S"),
                }
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)  # close section-card

    # ── Execution result panels ───────────────────────────────────────────────
    any_result = any(v for v in st.session_state.task_results.values())
    if any_result:
        st.markdown(
            '<div style="font-size:13px;font-weight:700;color:#0f172a;margin:20px 0 12px">'
            '▼ 執行結果</div>',
            unsafe_allow_html=True,
        )
        for task in MAIN_REPORT_TASKS:
            r = st.session_state.task_results.get(task["task_key"])
            if not r:
                continue
            rc = r["code"]
            rc_badge = badge(f"exit {rc}", "green" if rc == 0 else "red")
            st.markdown(
                f'<div class="exec-panel">'
                f'<div class="exec-panel-title">'
                f'▶ {r["name"]}'
                f'&emsp;<span style="font-size:12px;color:#64748b;font-weight:500">{r["ran_at"]}</span>'
                f'&emsp;{rc_badge}'
                f'</div>',
                unsafe_allow_html=True,
            )
            if rc != 0 and r["stderr"].strip():
                st.markdown('<div class="exec-label">STDERR（失敗原因）</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="log-box">{highlight_log(r["stderr"])}</div>', unsafe_allow_html=True)
            with st.expander(f"完整輸出 — {r['name']}"):
                if r["stdout"].strip():
                    st.markdown('<div class="exec-label">STDOUT</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="log-box">{highlight_log(r["stdout"])}</div>', unsafe_allow_html=True)
                if r["stderr"].strip():
                    st.markdown('<div class="exec-label">STDERR</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="log-box">{highlight_log(r["stderr"])}</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)


# ── 業績報表 ──────────────────────────────────────────────────────────────────

elif page == "業績報表":
    st.markdown(
        '<div class="page-header"><div class="page-title">業績報表</div><div class="page-subtitle">Latest Data · Send Later</div></div>',
        unsafe_allow_html=True,
    )

    result = None
    c1, c2, c3 = st.columns([1, 1, 1.5])
    with c1:
        update_btn = st.button("🔄 更新資料", use_container_width=True)
    with c2:
        send_btn = st.button("📧 寄送目前結果", use_container_width=True)
    with c3:
        reload_btn = st.button("📂 重新讀取已存資料", use_container_width=True)

    if update_btn:
        with st.spinner("更新資料中…"):
            result = generate_sales_report(send_email=False, persist_dashboard=True, trigger="dashboard")

    if reload_btn:
        st.rerun()

    # Load payload
    if result is not None:
        df4 = result.get("df4", pd.DataFrame())
        daily_df = result.get("daily_df", pd.DataFrame())
        email_html = result.get("email_html", "")
        updated_at = result.get("updated_at", "")
        exec_log_df = result.get("execution_log_df", pd.DataFrame())
        error_msg = result.get("error")
    else:
        payload = load_sales_latest_payload()
        df4 = payload.get("df4", pd.DataFrame())
        daily_df = payload.get("daily_df", pd.DataFrame())
        meta = payload.get("meta", {})
        email_html = payload.get("email_html", "")
        updated_at = meta.get("updated_at", "尚未產生資料") if isinstance(meta, dict) else "—"
        exec_log_df = load_execution_log_for_current_month()
        error_msg = meta.get("error") if isinstance(meta, dict) else None
        # Show load errors for debugging
        if payload.get("df4_error"):
            st.warning(f"df4.csv 讀取錯誤：{payload['df4_error']}")
        if payload.get("daily_df_error"):
            st.warning(f"daily_df.csv 讀取錯誤：{payload['daily_df_error']}")

    if send_btn:
        if df4.empty:
            st.warning("目前沒有可寄送資料，請先更新資料")
        else:
            try:
                from performance_report import send_region4_email
                send_region4_email(df4)
                st.success("寄信完成")
            except Exception as e:
                st.error(f"寄信失敗：{e}")

    if error_msg:
        st.error(f"上次執行有錯誤：{error_msg}")

    st.info(f"📅 最新更新時間：{updated_at}")

    # KPI metrics
    def _fmt_int(x):
        try: return f"{int(float(x)):,}"
        except: return "0"

    total = None
    if not df4.empty:
        t = df4[df4["城市"] == "加總"]
        if not t.empty:
            total = t.iloc[0]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("本月加總",     _fmt_int(total.get("本月加總", 0))     if total is not None else "—")
    k2.metric("次月加總",     _fmt_int(total.get("次月加總", 0))     if total is not None else "—")
    k3.metric("本月家電加總", _fmt_int(total.get("本月家電加總", 0)) if total is not None else "—")
    k4.metric("儲值金",       _fmt_int(total.get("儲值金", 0))       if total is not None else "—")

    # ── 各區月度摘要 (HTML table, right-aligned) ──────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📊 各區月度摘要</div>', unsafe_allow_html=True)

    if df4.empty:
        st.markdown('<div class="empty-state"><span class="icon">📭</span>目前沒有資料，請先按「更新資料」</div>', unsafe_allow_html=True)
    else:
        INT4  = {"本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"}
        PCT4  = {"本月佔比", "次月佔比"}
        RIGHT4 = INT4 | PCT4
        st.markdown(_html_table(df4, right_cols=RIGHT4, pct_cols=PCT4, int_cols=INT4), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── 當月每日業績總覽 (HTML table, right-aligned) ──────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📅 當月每日業績總覽</div>', unsafe_allow_html=True)

    # ── DEBUG: show state of daily_df ────────────────────────────────────────
    daily_dir = Path(LATEST_DIR)
    daily_csv = daily_dir / "daily_df.csv"
    debug_parts = []
    if daily_csv.exists():
        debug_parts.append(f"daily_df.csv 存在（{file_size_str(daily_csv)}，更新：{file_mtime(daily_csv)}）")
    else:
        debug_parts.append("⚠️ daily_df.csv 不存在")
    debug_parts.append(f"載入後筆數：{len(daily_df)} 行 × {len(daily_df.columns)} 欄")
    if not daily_df.empty:
        debug_parts.append(f"欄位：{', '.join(daily_df.columns.tolist())}")
    st.caption("  ·  ".join(debug_parts))
    # ── END DEBUG ─────────────────────────────────────────────────────────────

    if daily_df.empty:
        reason = ""
        if not daily_csv.exists():
            reason = "daily_df.csv 檔案不存在，請先按「更新資料」產生資料。"
        else:
            reason = "CSV 檔案存在但沒有資料列。可能原因：上次執行時日期欄解析失敗，或抓到的資料均無有效日期。請查看 Log 確認。"
        st.markdown(
            f'<div class="empty-state"><span class="icon">📭</span>{reason}</div>',
            unsafe_allow_html=True,
        )
    else:
        DAILY_COLS = [
            "日期",
            "台北業績", "台北佔比",
            "台中業績", "台中佔比",
            "桃園業績", "桃園佔比",
            "新竹業績", "新竹佔比",
            "高雄業績", "高雄佔比",
            "全區合計",
        ]
        exist_cols = [c for c in DAILY_COLS if c in daily_df.columns]
        df_daily_show = daily_df[exist_cols].copy()

        INT_D  = {c for c in exist_cols if "業績" in c or c == "全區合計"}
        PCT_D  = {c for c in exist_cols if "佔比" in c}
        RIGHT_D = INT_D | PCT_D

        st.markdown(_html_table(df_daily_show, right_cols=RIGHT_D, pct_cols=PCT_D, int_cols=INT_D), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── 當月累積執行紀錄 ──────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📝 當月累積執行紀錄</div>', unsafe_allow_html=True)

    if exec_log_df.empty:
        st.markdown('<div class="empty-state"><span class="icon">📋</span>目前沒有執行紀錄</div>', unsafe_allow_html=True)
    else:
        exec_ids = exec_log_df["id"].astype(str).tolist()
        sel_ids = st.multiselect("勾選要刪除的執行紀錄", options=exec_ids, key="del_exec")
        if st.button("🗑 刪除勾選列", key="del_exec_btn", use_container_width=True):
            deleted = delete_execution_log_rows(sel_ids)
            st.success(f"已刪除 {deleted} 筆")
            st.rerun()

        INT_E  = {"本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"}
        RIGHT_E = INT_E
        st.markdown(_html_table(exec_log_df, right_cols=RIGHT_E, pct_cols=set(), int_cols=INT_E), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── 信件預覽 ──────────────────────────────────────────────────────────────
    if email_html:
        with st.expander("📧 信件預覽"):
            st.components.v1.html(email_html, height=520, scrolling=True)


# ── 上下半月訂單 ──────────────────────────────────────────────────────────────

elif page == "上下半月訂單":
    st.markdown(
        '<div class="page-header"><div class="page-title">上下半月訂單</div><div class="page-subtitle">Half-Month Orders</div></div>',
        unsafe_allow_html=True,
    )
    st.info("這頁先保留。")


# ── 手動執行 ──────────────────────────────────────────────────────────────────

elif page == "手動執行":
    st.markdown(
        '<div class="page-header"><div class="page-title">手動執行</div><div class="page-subtitle">Manual Trigger</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">▶ 選擇任務執行</div>', unsafe_allow_html=True)
    selected = st.selectbox("選擇任務", MANUAL_TASKS, format_func=lambda x: x["name"])
    if st.button("▶ 執行", use_container_width=True):
        with st.spinner("執行中…"):
            rc, out, err = run_shell(selected["cmd"])
        rc_badge = badge(f"exit {rc}", "green" if rc == 0 else "red")
        st.markdown(f"回傳碼：{rc_badge}", unsafe_allow_html=True)
        if out.strip():
            st.markdown('<div class="exec-label">STDOUT</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="log-box">{highlight_log(out)}</div>', unsafe_allow_html=True)
        if err.strip():
            st.markdown('<div class="exec-label">STDERR</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="log-box">{highlight_log(err)}</div>', unsafe_allow_html=True)
        if not out.strip() and not err.strip():
            st.markdown('<div class="log-box"><span class="log-normal">(無輸出)</span></div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ── Log 監控 ──────────────────────────────────────────────────────────────────

elif page == "Log 監控":
    st.markdown(
        '<div class="page-header"><div class="page-title">Log 監控</div><div class="page-subtitle">Log Monitor</div></div>',
        unsafe_allow_html=True,
    )

    log_choices = {
        "主 log（cron.log）": LOG_FILE,
        "sales08 stderr": BASE_DIR / "launchd_sales08_stderr.log",
        "sales18 stderr": BASE_DIR / "launchd_sales18_stderr.log",
    }
    c1, c2, c3 = st.columns([3, 1, 1])
    with c1:
        sel_log = st.selectbox("選擇 log 檔", list(log_choices.keys()))
    with c2:
        n_lines = st.selectbox("顯示行數", [50, 100, 200, 500], index=1)
    with c3:
        if st.button("🔄 刷新", use_container_width=True):
            st.rerun()

    log_path = log_choices[sel_log]
    raw_log = read_last_lines(log_path, n_lines)
    st.markdown(f'<div class="log-meta">📄 {log_path}&emsp;·&emsp;更新：{file_mtime(log_path)}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="log-box">{highlight_log(raw_log)}</div>', unsafe_allow_html=True)


# ── 輸出檔案 ──────────────────────────────────────────────────────────────────

elif page == "輸出檔案":
    st.markdown(
        '<div class="page-header"><div class="page-title">輸出檔案監控</div><div class="page-subtitle">Output Files</div></div>',
        unsafe_allow_html=True,
    )

    rows = []
    for name, out_dir in OUTPUT_DIRS.items():
        files = find_latest_files(out_dir, limit=1)
        latest = files[0] if files else None
        rows.append({
            "分類": name,
            "最新檔案": latest.name if latest else "(無)",
            "時間": file_mtime(latest),
            "大小": file_size_str(latest),
            "資料夾": str(out_dir),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ── 程式管理 ──────────────────────────────────────────────────────────────────

elif page == "程式管理":
    st.markdown(
        '<div class="page-header"><div class="page-title">程式管理</div><div class="page-subtitle">Code Management</div></div>',
        unsafe_allow_html=True,
    )
    st.info("這頁先保留。")


# ── 排程設定 ──────────────────────────────────────────────────────────────────

elif page == "排程設定":
    st.markdown(
        '<div class="page-header"><div class="page-title">排程設定</div><div class="page-subtitle">Schedule Config</div></div>',
        unsafe_allow_html=True,
    )

    for task in MAIN_REPORT_TASKS:
        sched = load_plist_schedule(task["plist"], task["default_hour"], task["default_minute"])
        st.markdown(
            f'<div class="section-card">'
            f'<div class="section-title">⏰ {task["name"]}</div>',
            unsafe_allow_html=True,
        )
        st.code(task["cmd"], language="bash")

        col1, col2, col3, col4 = st.columns([1, 1, 1, 2])
        with col1:
            day = st.text_input("Day（空白=每日）", value=sched["day"], key=f"sday_{task['label']}")
        with col2:
            hour = st.text_input("Hour", value=sched["hour"], key=f"shour_{task['label']}")
        with col3:
            minute = st.text_input("Minute", value=sched["minute"], key=f"sminute_{task['label']}")
        with col4:
            st.markdown(
                f'<div class="next-run" style="margin-top:28px">🕐 下次執行：{calc_next_run(sched["day"], hour, minute)}</div>',
                unsafe_allow_html=True,
            )

        if st.button(f"儲存 {task['name']}", key=f"ssave_{task['label']}", use_container_width=True):
            if not IS_LOCAL:
                st.warning("雲端環境無法修改本機 plist")
            elif not task["plist"].exists():
                st.error(f"找不到 plist：{task['plist'].name}")
            elif not hour.isdigit() or not minute.isdigit():
                st.error("Hour / Minute 必須是數字")
            elif day.strip() and not day.isdigit():
                st.error("Day 必須留空或填數字")
            else:
                code, out, err = save_plist_schedule(task["plist"], hour, minute, day)
                st.success(f"✓ 已更新 {task['name']}") if code == 0 else st.error(err or "更新失敗")

        st.markdown("</div>", unsafe_allow_html=True)


# ── Footer ────────────────────────────────────────────────────────────────────

st.markdown(
    f'<div class="footer-cap">Lemon Clean Scheduler Console &nbsp;·&nbsp; {now_taipei().strftime("%Y-%m-%d %H:%M:%S")}</div>',
    unsafe_allow_html=True,
)
