import json
import math
import os
import plistlib
import py_compile
import shutil
import subprocess
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Callable, Dict, List

import pandas as pd
import streamlit as st

from paths import (
    PATH_CLEANER_DATA,
    PATH_CLEANER_SCHEDULE,
    PATH_HR,
    PATH_JENNY,
    PATH_ORDER,
    PATH_SCHEDULE,
    PATH_VIP,
)

from 業績報表 import (
    generate_sales_report,
    load_execution_log_for_current_month,
    load_daily_history_for_current_month,
    delete_daily_history_rows,
    LATEST_DIR,
)

# ══════════════════════════════════════════════════
# 常數設定
# ══════════════════════════════════════════════════
BASE_DIR = Path("/Users/jenny/lemon")
LOG_FILE = BASE_DIR / "cron.log"
LAUNCH_AGENTS_DIR = Path.home() / "Library/LaunchAgents"
CONFIG_FILE = BASE_DIR / "dashboard_config.json"

DEFAULT_CONFIG = {
    "yyyymm_scripts": ["預收.py", "已退款.py"],
    "halfmonth_scripts": ["上下半月訂單.py"],
}

OUTPUT_DIRS = {
    "排班統計表": Path(PATH_SCHEDULE),
    "專員班表": Path(PATH_CLEANER_SCHEDULE),
    "專員系統個資": Path(PATH_CLEANER_DATA),
    "訂單資料": Path(PATH_ORDER),
    "業績報表": Path(f"{PATH_JENNY}/業績報表"),
    "預收": Path(PATH_VIP),
    "儲值金結算": Path(PATH_VIP),
    "儲值金預收": Path(PATH_VIP),
    "上下半月訂單": Path(PATH_HR),
    "已退款": Path(PATH_HR),
}

CATEGORY_MATCHERS: Dict[str, Callable[[Path], bool]] = {
    "排班統計表": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "專員班表": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "專員系統個資": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "訂單資料": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "業績報表": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".html", ".json"},
    "預收": lambda p: "預收" in p.name and "儲值金預收" not in p.name and p.suffix.lower() in {".xlsx", ".xls", ".csv"},
    "儲值金結算": lambda p: "儲值金結算" in p.name and p.suffix.lower() in {".xlsx", ".xls", ".csv"},
    "儲值金預收": lambda p: "儲值金預收" in p.name and p.suffix.lower() in {".xlsx", ".xls", ".csv"},
    "上下半月訂單": lambda p: "訂單-" in p.name and "已退款" not in p.name and p.suffix.lower() in {".xlsx", ".xls", ".csv"},
    "已退款": lambda p: "已退款" in p.name and p.suffix.lower() in {".xlsx", ".xls", ".csv"},
}

TASK_OUTPUT_MAP = {
    "com.jenny.daily01": ["排班統計表", "專員班表"],
    "com.jenny.daily02": ["訂單資料", "專員系統個資"],
    "com.jenny.sales08": ["業績報表"],
    "com.jenny.sales18": ["業績報表"],
    "com.jenny.monthlyfirst": ["預收", "儲值金結算", "儲值金預收"],
    "com.jenny.midmonth": ["上下半月訂單"],
    "com.jenny.monthend": ["上下半月訂單", "已退款"],
}

SCHEDULE_TASKS = [
    {
        "name": "每日 01:00",
        "label": "com.jenny.daily01",
        "desc": "排班統計表 + 專員班表",
        "cmd": f'cd "{BASE_DIR}" && bash "{BASE_DIR}/launchd_jobs/daily_01.sh"',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01.plist",
        "hour": 1,
        "minute": 0,
        "day": None,
    },
    {
        "name": "每日 02:00",
        "label": "com.jenny.daily02",
        "desc": "當月次月訂單 + 專員系統個資",
        "cmd": f'cd "{BASE_DIR}" && bash "{BASE_DIR}/launchd_jobs/daily_02.sh"',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02.plist",
        "hour": 2,
        "minute": 0,
        "day": None,
    },
    {
        "name": "每日 08:00",
        "label": "com.jenny.sales08",
        "desc": "業績報表",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 業績報表.py schedule true',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.sales08.plist",
        "hour": 8,
        "minute": 0,
        "day": None,
    },
    {
        "name": "每日 18:00",
        "label": "com.jenny.sales18",
        "desc": "業績報表",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 業績報表.py schedule true',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.sales18.plist",
        "hour": 18,
        "minute": 0,
        "day": None,
    },
    {
        "name": "每月 1 日",
        "label": "com.jenny.monthlyfirst",
        "desc": "預收 + 儲值金結算 + 儲值金預收",
        "cmd": f'cd "{BASE_DIR}" && bash "{BASE_DIR}/launchd_jobs/monthly_first.sh"',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.monthlyfirst.plist",
        "hour": 1,
        "minute": 0,
        "day": 1,
    },
    {
        "name": "每月 15 日",
        "label": "com.jenny.midmonth",
        "desc": "上半月訂單",
        "cmd": f'cd "{BASE_DIR}" && bash "{BASE_DIR}/launchd_jobs/monthly_15.sh"',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.midmonth.plist",
        "hour": 1,
        "minute": 0,
        "day": 15,
    },
    {
        "name": "每月月底",
        "label": "com.jenny.monthend",
        "desc": "下半月訂單 + 已退款",
        "cmd": f'cd "{BASE_DIR}" && bash "{BASE_DIR}/launchd_jobs/month_end.sh"',
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.monthend.plist",
        "hour": 1,
        "minute": 0,
        "day": 28,
    },
]

LAUNCHD_STDERR_MAP = {
    "com.jenny.daily01": BASE_DIR / "launchd_daily01_stderr.log",
    "com.jenny.daily02": BASE_DIR / "launchd_daily02_stderr.log",
    "com.jenny.sales08": BASE_DIR / "launchd_sales08_stderr.log",
    "com.jenny.sales18": BASE_DIR / "launchd_sales18_stderr.log",
    "com.jenny.monthlyfirst": BASE_DIR / "launchd_monthlyfirst_stderr.log",
    "com.jenny.midmonth": BASE_DIR / "launchd_midmonth_stderr.log",
    "com.jenny.monthend": BASE_DIR / "launchd_monthend_stderr.log",
}

NAV_PAGES = ["主控表", "手動", "Log", "輸出", "腳本", "排程", "業績"]
NAV_ICONS = ["📋", "▶️", "📄", "📂", "⚙️", "⏰", "📊"]

# ══════════════════════════════════════════════════
# 工具函式
# ══════════════════════════════════════════════════
def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def run_shell(cmd: str) -> tuple[int, str, str]:
    p = subprocess.run(cmd, shell=True, text=True, capture_output=True, executable="/bin/bash")
    return p.returncode, p.stdout, p.stderr


def get_launchd_status() -> dict:
    code, out, _ = run_shell("launchctl list | grep com.jenny")
    status_map = {}
    if code != 0 and not out.strip():
        return status_map
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 3:
            pid, last_exit, label = parts[0], parts[1], parts[2]
            status_map[label] = {"pid": pid, "last_exit": last_exit}
    return status_map


def read_last_lines(path: Path, n: int = 120) -> str:
    if not path.exists():
        return "(尚無 log)"
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        return "".join(lines[-n:])
    except Exception as e:
        return f"(讀取失敗) {e}"


def file_mtime(path: Optional[Path]) -> str:
    if not path or not path.exists():
        return "-"
    ts = datetime.fromtimestamp(path.stat().st_mtime)
    return ts.strftime("%m/%d %H:%M")


def file_mtime_dt(path: Optional[Path]):
    if not path or not path.exists():
        return None
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return None


def format_dt(ts) -> str:
    return "—" if not ts else ts.strftime("%m/%d %H:%M")


def file_size_str(path: Optional[Path]) -> str:
    if not path or not path.exists():
        return "-"
    size = path.stat().st_size
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size/1024:.1f} KB"
    return f"{size/1024/1024:.1f} MB"


def render_status_info(label: str, status_map: dict) -> dict:
    info = status_map.get(label)
    if not info:
        return {"label": "未載入", "cls": "gray"}
    pid, last_exit = info["pid"], info["last_exit"]
    if pid != "-":
        return {"label": f"執行中 PID {pid}", "cls": "yellow"}
    if last_exit == "0":
        return {"label": "正常", "cls": "green"}
    return {"label": f"異常 exit {last_exit}", "cls": "red"}


def today_status_info(log_text: str) -> dict:
    today1 = datetime.today().strftime("%Y-%m-%d")
    today2 = datetime.today().strftime("%Y%m%d")
    if today1 not in log_text and today2 not in log_text:
        return {"label": "今日未執行", "cls": "gray"}
    if "Traceback" in log_text or "❌" in log_text or "PermissionError" in log_text:
        return {"label": "今日有錯誤", "cls": "red"}
    if "✅" in log_text:
        return {"label": "今日成功", "cls": "green"}
    return {"label": "今日有執行", "cls": "yellow"}


def list_python_files():
    return sorted(BASE_DIR.glob("*.py"), key=lambda p: p.name.lower())


def classify_py(name: str) -> str:
    if "報表" in name:
        return "報表"
    if "排班" in name or "排程" in name:
        return "排程"
    if "test" in name.lower() or "測試" in name:
        return "測試"
    return "其他"


def is_valid_output_file(path: Path, category: str) -> bool:
    if not path.is_file():
        return False
    if path.name.startswith(".") or path.name.startswith("~$"):
        return False
    matcher = CATEGORY_MATCHERS.get(category)
    if matcher is None:
        return True
    try:
        return bool(matcher(path))
    except Exception:
        return False


def find_latest_files(base_dir: Path, limit: int = 10, category: Optional[str] = None):
    if not base_dir.exists():
        return []
    files: List[Path] = []
    try:
        for p in base_dir.rglob("*"):
            if not p.is_file():
                continue
            if p.name.startswith(".") or p.name.startswith("~$"):
                continue
            if category and not is_valid_output_file(p, category):
                continue
            files.append(p)
    except Exception:
        return []
    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return files[:limit]


def get_latest_output_file(category: str):
    out_dir = OUTPUT_DIRS.get(category)
    if not out_dir:
        return None
    files = find_latest_files(out_dir, limit=1, category=category)
    return files[0] if files else None


def get_output_completion_info(category: str) -> dict:
    latest_file = get_latest_output_file(category)
    latest_dt = file_mtime_dt(latest_file)
    latest_is_today = bool(latest_dt and latest_dt.date() == datetime.today().date())

    return {
        "category": category,
        "file": latest_file,
        "dt": latest_dt if latest_is_today else None,
        "is_complete": bool(latest_file and latest_is_today),
        "latest_any_dt": latest_dt,
    }


def get_task_output_summary(task: dict) -> dict:
    categories = TASK_OUTPUT_MAP.get(task["label"], [])
    if not categories:
        return {
            "required_count": 0,
            "complete_count": 0,
            "all_complete": False,
            "any_complete": False,
            "latest_complete_dt": None,
            "details": [],
        }

    details = [get_output_completion_info(category) for category in categories]
    complete_details = [d for d in details if d["is_complete"]]
    latest_complete_dt = max([d["dt"] for d in complete_details], default=None)

    return {
        "required_count": len(categories),
        "complete_count": len(complete_details),
        "all_complete": len(complete_details) == len(categories),
        "any_complete": len(complete_details) > 0,
        "latest_complete_dt": latest_complete_dt,
        "details": details,
    }


def get_task_update_dt(task: dict):
    summary = get_task_output_summary(task)
    if summary["all_complete"]:
        return summary["latest_complete_dt"]
    return None


def load_plist_schedule(plist_path: Path):
    if not plist_path.exists():
        return {"type": "missing", "day": "", "hour": "", "minute": ""}
    try:
        with open(plist_path, "rb") as f:
            data = plistlib.load(f)
        interval = data.get("StartCalendarInterval")
        if isinstance(interval, list):
            return {"type": "list", "day": "", "hour": "", "minute": ""}
        if isinstance(interval, dict):
            return {
                "type": "dict",
                "day": str(interval.get("Day", "")),
                "hour": str(interval.get("Hour", "")),
                "minute": str(interval.get("Minute", "")),
            }
        return {"type": "unknown", "day": "", "hour": "", "minute": ""}
    except Exception:
        return {"type": "error", "day": "", "hour": "", "minute": ""}


def save_plist_schedule(plist_path: Path, day: str, hour: str, minute: str):
    with open(plist_path, "rb") as f:
        data = plistlib.load(f)
    interval = {}
    if day.strip():
        interval["Day"] = int(day)
    interval["Hour"] = int(hour)
    interval["Minute"] = int(minute)
    data["StartCalendarInterval"] = interval
    with open(plist_path, "wb") as f:
        plistlib.dump(data, f)
    label = data.get("Label")
    run_shell(f'launchctl bootout gui/$(id -u) "{plist_path}" 2>/dev/null')
    code, out, err = run_shell(f'launchctl bootstrap gui/$(id -u) "{plist_path}"')
    return label, code, out, err


def calc_next_run(day_str: str, hour_str: str, minute_str: str) -> str:
    try:
        hour = int(hour_str) if hour_str else 0
        minute = int(minute_str) if minute_str else 0
        now = datetime.now()

        if day_str.strip():
            day = int(day_str)
            candidate = now.replace(day=day, hour=hour, minute=minute, second=0, microsecond=0)
            if candidate <= now:
                if now.month == 12:
                    candidate = candidate.replace(year=now.year + 1, month=1)
                else:
                    candidate = candidate.replace(month=now.month + 1)
            return candidate.strftime("%Y-%m-%d  %H:%M")

        candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate <= now:
            candidate += timedelta(days=1)
        return candidate.strftime("%Y-%m-%d  %H:%M")
    except Exception:
        return "—"


def highlight_log(text: str) -> str:
    html_lines = []
    for line in text.splitlines():
        escaped = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if any(k in line for k in ["Traceback", "Error", "ERROR", "❌", "PermissionError", "FAILED", "failed"]):
            html_lines.append(f'<span class="log-err">{escaped}</span>')
        elif any(k in line for k in ["✅", "SUCCESS", "success", "完成", "Done", "done"]):
            html_lines.append(f'<span class="log-ok">{escaped}</span>')
        elif any(k in line for k in ["WARNING", "Warning", "warn", "⚠"]):
            html_lines.append(f'<span class="log-warn">{escaped}</span>')
        elif any(k in line for k in ["INFO", "info", "開始", "Start", "start"]):
            html_lines.append(f'<span class="log-info">{escaped}</span>')
        else:
            html_lines.append(f'<span class="log-normal">{escaped}</span>')
    return "\n".join(html_lines)


def open_in_finder(path: Path):
    try:
        if path.is_file():
            subprocess.Popen(["open", "-R", str(path)])
        else:
            subprocess.Popen(["open", str(path)])
    except Exception as e:
        st.error(f"無法開啟 Finder：{e}")


# ─────────────────────────────────────────
# 業績報表工具
# ─────────────────────────────────────────
def load_sales_latest_payload():
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

    if os.path.exists(df4_path):
        payload["df4"] = pd.read_csv(df4_path, encoding="utf-8-sig")
    if os.path.exists(daily_path):
        payload["daily_df"] = pd.read_csv(daily_path, encoding="utf-8-sig")
    if os.path.exists(meta_path):
        with open(meta_path, "r", encoding="utf-8") as f:
            payload["meta"] = json.load(f)
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            payload["email_html"] = f.read()

    return payload


def sales_fmt_int(x):
    try:
        return f"{int(float(x)):,}"
    except Exception:
        return "0"


def sales_fmt_pct(x):
    try:
        return f"{float(x):.2%}"
    except Exception:
        return "0.00%"


def style_sales_df4(df):
    if df.empty:
        return df
    out = df.copy()
    for col in ["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"]:
        if col in out.columns:
            out[col] = out[col].apply(sales_fmt_int)
    for col in ["本月佔比", "次月佔比"]:
        if col in out.columns:
            out[col] = out[col].apply(sales_fmt_pct)
    return out


def style_sales_daily(df):
    if df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if col.endswith("業績") or col == "全區合計":
            out[col] = out[col].apply(sales_fmt_int)
        elif col.endswith("佔比"):
            out[col] = out[col].apply(sales_fmt_pct)
    return out


def style_sales_exec(df):
    if df.empty:
        return df
    out = df.copy()
    for col in ["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"]:
        if col in out.columns:
            out[col] = out[col].apply(sales_fmt_int)
    return out


def style_sales_history(df):
    if df.empty:
        return df
    out = df.copy()
    if "今日全區合計" in out.columns:
        out["今日全區合計"] = out["今日全區合計"].apply(sales_fmt_int)
    return out


def get_sales_total_row(df4):
    if df4.empty:
        return None
    total = df4[df4["城市"] == "加總"]
    if total.empty:
        return None
    return total.iloc[0]


def sales_paginate_df(df, page_key, page_size=10):
    if df.empty:
        return df, 1, 1, 0

    total_rows = len(df)
    total_pages = max(1, math.ceil(total_rows / page_size))

    if page_key not in st.session_state:
        st.session_state[page_key] = 1

    st.session_state[page_key] = max(1, min(st.session_state[page_key], total_pages))
    page = st.session_state[page_key]

    start = (page - 1) * page_size
    end = start + page_size

    return df.iloc[start:end].copy(), page, total_pages, total_rows


# ══════════════════════════════════════════════════
# 頁面設定
# ══════════════════════════════════════════════════
st.set_page_config(page_title="營運報表控制台", page_icon="📊", layout="wide")

if "page" not in st.session_state:
    st.session_state.page = "排程主控表"
if "sales_exec_page" not in st.session_state:
    st.session_state.sales_exec_page = 1
if "sales_history_page" not in st.session_state:
    st.session_state.sales_history_page = 1
if "sales_delete_ids" not in st.session_state:
    st.session_state.sales_delete_ids = []

# ══════════════════════════════════════════════════
# 全站樣式
# ══════════════════════════════════════════════════
st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] { background: #f3f4f6 !important; font-family: sans-serif; color: #1c2333; }
[data-testid="stHeader"], [data-testid="stSidebar"] { display: none !important; }
.block-container { padding: 0 2.4rem 3rem !important; max-width: 1800px !important; }
.topbar { background: #fff; border: 1px solid #e5e7eb; border-radius: 18px; padding: 24px 28px; margin: 24px 0 28px; display:flex; align-items:center; justify-content:space-between; box-shadow: 0 4px 14px rgba(0,0,0,0.05); }
.brand { display:flex; align-items:center; gap:18px; }
.brand-title { font-size: 28px; font-weight: 800; color:#0f766e; }
.brand-badge { border:1px solid #cbd5e1; border-radius: 999px; padding: 10px 18px; color:#94a3b8; font-weight:700; font-size:18px; }
.clock { font-family: monospace; font-size: 20px; font-weight: 700; color:#475569; }
.page-tabs button { height:110px !important; font-size:22px !important; font-weight:800 !important; border-radius:18px !important; border:3px solid #cbd5e1 !important; background:#fff !important; color:#0f172a !important; }
.page-tabs .active button { border-color:#60a5fa !important; background:#eff6ff !important; }
.panel { background: #ffffff; border: 1px solid #e5e7eb; border-radius: 18px; padding: 24px; margin-bottom: 20px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); }
.panel-title { font-size: 18px; font-weight: 800; color:#0f172a; margin-bottom:14px; }
.page-title { font-size: 42px; font-weight: 900; color:#0f172a; margin-bottom:8px; }
.page-subtitle { font-size: 18px; color:#94a3b8; margin-bottom:20px; letter-spacing: 0.08em; font-family: monospace; }
.log-box { background: #0d1117; border-radius: 12px; padding: 16px 20px; font-family: monospace; font-size: 12.5px; line-height: 1.7; white-space: pre-wrap; word-break: break-all; max-height: 500px; overflow: auto; border: 1px solid #1e2a3a; }
.log-err { color: #f87171; display: block; }
.log-ok { color: #4ade80; display: block; }
.log-warn { color: #fbbf24; display: block; }
.log-info { color: #60a5fa; display: block; }
.log-normal { color: #8b9ab0; display: block; }
.badge { display: inline-flex; align-items: center; gap: 5px; font-size: 11.5px; font-weight: 600; padding: 3px 10px; border-radius: 20px; white-space: nowrap; }
.badge .bdot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
.badge.green { background: rgba(5,150,105,0.1); border: 1px solid rgba(5,150,105,0.25); color: #065f46; }
.badge.green .bdot { background: #10b981; }
.badge.yellow { background: rgba(217,119,6,0.1); border: 1px solid rgba(217,119,6,0.25); color: #92400e; }
.badge.yellow .bdot { background: #f59e0b; }
.badge.red { background: rgba(220,38,38,0.1); border: 1px solid rgba(220,38,38,0.25); color: #991b1b; }
.badge.red .bdot { background: #ef4444; }
.badge.gray { background: rgba(148,163,184,0.1); border: 1px solid rgba(148,163,184,0.25); color: #64748b; }
.badge.gray .bdot { background: #94a3b8; }
.next-run-box { background: rgba(29,78,216,0.06); border: 1px solid rgba(29,78,216,0.15); border-radius: 10px; padding: 10px 16px; margin-top: 10px; font-size: 12.5px; color: #1d4ed8; display: flex; align-items: center; gap: 8px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# 頂部
# ══════════════════════════════════════════════════
now_str = datetime.now().strftime("%Y/%m/%d %H:%M 台北")
st.markdown(
    f"""
    <div class="topbar">
        <div class="brand">
            <div class="brand-title">📊 營運報表控制台</div>
            <div class="brand-badge">📍 未選地區</div>
        </div>
        <div class="clock">🕒 {now_str}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_cols = st.columns(len(NAV_PAGES))
for i, (label, icon) in enumerate(zip(NAV_PAGES, NAV_ICONS)):
    with tab_cols[i]:
        if st.button(f"{icon} {label}", key=f"tab_{label}", use_container_width=True):
            st.session_state.page = label
            st.rerun()

st.divider()

page = st.session_state.page
cfg = load_config()
status_map = get_launchd_status()

# ══════════════════════════════════════════════════
# 排程主控表
# ══════════════════════════════════════════════════
if page == "排程主控表":
    st.markdown('<div class="page-title">排程主控表</div><div class="page-subtitle">SCHEDULE DASHBOARD</div>', unsafe_allow_html=True)
    main_log_text = read_last_lines(LOG_FILE, 300)
    task_data = []

    for task in SCHEDULE_TASKS:
        stderr_path = LAUNCHD_STDERR_MAP.get(task["label"])
        stderr_text = read_last_lines(stderr_path, 100) if stderr_path else ""
        sched = load_plist_schedule(task["plist"])
        output_summary = get_task_output_summary(task)

        if sched["type"] == "dict":
            sched_text = (
                f'每月 {sched["day"]} 日  {sched["hour"].zfill(2)}:{sched["minute"].zfill(2)}'
                if sched["day"]
                else f'每日 {sched["hour"].zfill(2)}:{sched["minute"].zfill(2)}'
            )
        else:
            sched_text = "—"

        launchd = render_status_info(task["label"], status_map)
        stderr_today = today_status_info(stderr_text) if stderr_text.strip() else {"label": "今日未執行", "cls": "gray"}
        main_today = today_status_info(main_log_text)

        if stderr_today["cls"] == "red":
            today = {"label": "今日有錯誤", "cls": "red"}
        elif launchd["cls"] == "yellow":
            today = {"label": "執行中", "cls": "yellow"}
        elif output_summary["all_complete"]:
            today = {"label": "今日成功", "cls": "green"}
        elif output_summary["any_complete"]:
            today = {"label": "部分完成", "cls": "yellow"}
        elif stderr_today["cls"] == "green" or main_today["cls"] != "gray":
            today = {"label": "已執行待確認", "cls": "yellow"}
        else:
            today = {"label": "今日未執行", "cls": "gray"}

        task_data.append({
            "task": task,
            "sched_text": sched_text,
            "launchd": launchd,
            "today": today,
            "done_text": f'{output_summary["complete_count"]}/{output_summary["required_count"]}',
            "done_dt": format_dt(get_task_update_dt(task)),
        })

    st.markdown('<div class="panel"><div class="panel-title">任務狀態</div>', unsafe_allow_html=True)

    selected_batch_tasks = []

    for d in task_data:
        task = d["task"]
        lk = d["launchd"]
        tk = d["today"]

        row_cols = st.columns([0.4, 1.2, 2.5, 1.2, 2.0, 1.0, 0.8])

        with row_cols[0]:
            checked = st.checkbox("選取任務", key=f"batch_{task['label']}", label_visibility="collapsed")
            if checked:
                selected_batch_tasks.append(task)

        with row_cols[1]:
            st.markdown(f"**{task['name']}**")

        with row_cols[2]:
            st.caption(f"{task['desc']} · 輸出完成 {d['done_text']}")

        with row_cols[3]:
            st.caption(d["sched_text"])

        with row_cols[4]:
            st.markdown(
                f'<span class="badge {lk["cls"]}"><span class="bdot"></span>{lk["label"]}</span> '
                f'<span class="badge {tk["cls"]}"><span class="bdot"></span>{tk["label"]}</span>',
                unsafe_allow_html=True,
            )

        with row_cols[5]:
            st.caption(d["done_dt"])

        with row_cols[6]:
            if st.button("▶", key=f"run_main_{task['label']}", use_container_width=True):
                with st.spinner(f'執行中：{task["name"]}'):
                    code, out, err = run_shell(task["cmd"])
                st.write(f"### {task['name']} 執行結果")
                st.write(f"回傳碼：`{code}`")
                st.text_area(f"stdout_{task['label']}", out or "(無輸出)", height=220)
                st.text_area(f"stderr_{task['label']}", err or "(無錯誤)", height=180)

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("🔄 重新整理狀態", use_container_width=True):
            st.rerun()
    with c2:
        if st.button("♻️ 重新載入全部排程", use_container_width=True):
            code, out, err = run_shell(
                'for f in ~/Library/LaunchAgents/com.jenny*.plist; do '
                'launchctl bootout gui/$(id -u) "$f" 2>/dev/null; '
                'launchctl bootstrap gui/$(id -u) "$f"; '
                'done'
            )
            st.code((out or "") + ("\n" + err if err else ""))
            st.rerun()
    with c3:
        if st.button("▶ 執行勾選項目", use_container_width=True):
            if not selected_batch_tasks:
                st.warning("請先勾選至少一個任務")
            else:
                batch_result_rows = []
                for task in selected_batch_tasks:
                    with st.spinner(f'批次執行中：{task["name"]}'):
                        code, out, err = run_shell(task["cmd"])
                    batch_result_rows.append({
                        "任務": task["name"],
                        "回傳碼": code,
                        "結果": "成功" if code == 0 else "失敗",
                    })
                st.dataframe(pd.DataFrame(batch_result_rows), use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# 手動執行
# ══════════════════════════════════════════════════
elif page == "手動執行":
    st.markdown('<div class="page-title">手動執行</div><div class="page-subtitle">MANUAL TRIGGER</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">排程任務</div>', unsafe_allow_html=True)
    schedule_selected = st.selectbox("選擇排程任務", SCHEDULE_TASKS, format_func=lambda x: f'{x["name"]} · {x["desc"]}', key="schedule_task_select")
    if st.button(f'▶ 執行：{schedule_selected["name"]}', use_container_width=True):
        with st.spinner("執行中…"):
            code, out, err = run_shell(schedule_selected["cmd"])
        st.write(f"回傳碼：`{code}`")
        st.text_area("stdout", out or "(無輸出)", height=220)
        st.text_area("stderr", err or "(無錯誤)", height=180)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">單支 .py</div>', unsafe_allow_html=True)
    py_files = list_python_files()
    py_selected = st.selectbox("選擇 Python 腳本", py_files, format_func=lambda x: x.name, key="py_task_select")
    month_arg = ""
    half_arg = ""
    if py_selected.name in cfg.get("yyyymm_scripts", []) or py_selected.name in cfg.get("halfmonth_scripts", []):
        month_arg = st.text_input("輸入月份（YYYYMM）", value=datetime.today().strftime("%Y%m"), key="month_arg_input")
    if py_selected.name in cfg.get("halfmonth_scripts", []):
        half_arg = st.selectbox("選擇半月", ["1", "2"], format_func=lambda x: "上半月" if x == "1" else "下半月", key="half_arg_select")

    if st.button(f'▶ 執行：{py_selected.name}', use_container_width=True):
        if py_selected.name in cfg.get("halfmonth_scripts", []):
            if len(month_arg) != 6 or not month_arg.isdigit():
                st.error("月份格式請輸入 YYYYMM，例如 202603")
            else:
                cmd = f'cd "{BASE_DIR}" && /usr/bin/python3 "{py_selected.name}" {month_arg} {half_arg}'
                code, out, err = run_shell(cmd)
                st.write(f"回傳碼：`{code}`")
                st.text_area("stdout", out or "(無輸出)", height=220)
                st.text_area("stderr", err or "(無錯誤)", height=180)
        elif py_selected.name in cfg.get("yyyymm_scripts", []):
            if len(month_arg) != 6 or not month_arg.isdigit():
                st.error("月份格式請輸入 YYYYMM，例如 202603")
            else:
                cmd = f'cd "{BASE_DIR}" && /usr/bin/python3 "{py_selected.name}" {month_arg}'
                code, out, err = run_shell(cmd)
                st.write(f"回傳碼：`{code}`")
                st.text_area("stdout", out or "(無輸出)", height=220)
                st.text_area("stderr", err or "(無錯誤)", height=180)
        else:
            cmd = f'cd "{BASE_DIR}" && /usr/bin/python3 "{py_selected.name}"'
            code, out, err = run_shell(cmd)
            st.write(f"回傳碼：`{code}`")
            st.text_area("stdout", out or "(無輸出)", height=220)
            st.text_area("stderr", err or "(無錯誤)", height=180)
    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# Log 監控
# ══════════════════════════════════════════════════
elif page == "Log 監控":
    st.markdown('<div class="page-title">Log 監控</div><div class="page-subtitle">LOG MONITOR</div>', unsafe_allow_html=True)
    st.markdown('<div class="panel"><div class="panel-title">Log 檔案</div>', unsafe_allow_html=True)

    log_choices = {
        "主 log（cron.log）": LOG_FILE,
        "daily01 stderr": BASE_DIR / "launchd_daily01_stderr.log",
        "daily02 stderr": BASE_DIR / "launchd_daily02_stderr.log",
        "sales08 stderr": BASE_DIR / "launchd_sales08_stderr.log",
        "sales18 stderr": BASE_DIR / "launchd_sales18_stderr.log",
        "monthlyfirst stderr": BASE_DIR / "launchd_monthlyfirst_stderr.log",
        "midmonth stderr": BASE_DIR / "launchd_midmonth_stderr.log",
        "monthend stderr": BASE_DIR / "launchd_monthend_stderr.log",
    }

    c1, c2, c3 = st.columns([3, 1, 1])
    with c1:
        selected_log = st.selectbox("選擇 log 檔", list(log_choices.keys()))
    with c2:
        n_lines = st.selectbox("顯示行數", [50, 100, 200, 500], index=1)
    with c3:
        if st.button("🔄 手動刷新", use_container_width=True):
            st.rerun()

    log_path = log_choices[selected_log]
    raw_log = read_last_lines(log_path, n_lines)
    st.caption(f"檔案：{log_path} · 更新：{file_mtime(log_path)}")
    st.markdown(f'<div class="log-box">{highlight_log(raw_log)}</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# 輸出報表
# ══════════════════════════════════════════════════
elif page == "輸出報表":
    st.markdown('<div class="page-title">輸出報表</div><div class="page-subtitle">OUTPUT FILES</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">總覽</div>', unsafe_allow_html=True)
    rows = []
    for name, out_dir in OUTPUT_DIRS.items():
        info = get_output_completion_info(name)
        latest_any_file = info["file"]
        rows.append({
            "分類": name,
            "最新檔案": latest_any_file.name if latest_any_file else "(無)",
            "狀態": "今日完成" if info["is_complete"] else "未完成",
            "大小": file_size_str(latest_any_file) if latest_any_file else "—",
            "完成時間": format_dt(info["dt"]),
            "最新檔時間": format_dt(info["latest_any_dt"]),
            "資料夾": str(out_dir),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">詳細清單</div>', unsafe_allow_html=True)
    selected_folder = st.selectbox("查看哪個資料夾", list(OUTPUT_DIRS.keys()))
    c1, c2 = st.columns([3, 1])
    with c2:
        if st.button("📁 在 Finder 開啟", use_container_width=True):
            open_in_finder(OUTPUT_DIRS[selected_folder])
            st.success("已開啟 Finder")

    search_keyword = st.text_input("搜尋檔名", value="", key="output_search_keyword")
    files = find_latest_files(OUTPUT_DIRS[selected_folder], limit=100, category=selected_folder)

    file_rows = []
    for f in files:
        if search_keyword and search_keyword.lower() not in f.name.lower():
            continue
        dt = file_mtime_dt(f)
        is_complete = bool(dt and dt.date() == datetime.today().date())
        file_rows.append({
            "檔名": f.name,
            "狀態": "今日完成" if is_complete else "非今日 / 僅供參考",
            "大小": file_size_str(f),
            "完成時間": format_dt(dt) if is_complete else "—",
            "最新檔時間": file_mtime(f),
            "完整路徑": str(f),
        })

    if file_rows:
        st.dataframe(pd.DataFrame(file_rows), use_container_width=True, hide_index=True)
    else:
        st.info("此資料夾沒有符合條件的檔案")
    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# 腳本管理
# ══════════════════════════════════════════════════
elif page == "腳本管理":
    st.markdown('<div class="page-title">腳本管理</div><div class="page-subtitle">CODE MANAGEMENT</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">新增 .py</div>', unsafe_allow_html=True)
    new_filename = st.text_input("檔名（例如 test.py）", key="new_py_filename")
    new_content = st.text_area("程式內容", value="# new python file\n\nprint('hello')\n", height=180, key="new_py_content")
    if st.button("建立 .py"):
        if not new_filename.endswith(".py"):
            st.error("檔名必須以 .py 結尾")
        else:
            new_path = BASE_DIR / new_filename
            if new_path.exists():
                st.error("檔案已存在")
            else:
                new_path.write_text(new_content, encoding="utf-8")
                st.success(f"已建立：{new_filename}")
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">執行參數設定</div>', unsafe_allow_html=True)
    py_files = list_python_files()
    py_names = [p.name for p in py_files]
    selected_yyyymm = st.multiselect("需要輸入 YYYYMM 的 .py", py_names, default=cfg.get("yyyymm_scripts", []))
    selected_halfmonth = st.multiselect("需要輸入 YYYYMM + 半月 的 .py", py_names, default=cfg.get("halfmonth_scripts", []))
    if st.button("儲存參數設定"):
        cfg["yyyymm_scripts"] = selected_yyyymm
        cfg["halfmonth_scripts"] = selected_halfmonth
        save_config(cfg)
        st.success("已儲存設定")
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">修改 / 刪除 / 測試</div>', unsafe_allow_html=True)
    if not py_files:
        st.warning("找不到任何 .py 檔")
    else:
        selected_file = st.selectbox(
            "選擇要管理的 .py",
            py_files,
            format_func=lambda p: f"{p.name} · {classify_py(p.name)} · {file_size_str(p)}",
            key="edit_py_select",
        )
        selected_path = str(selected_file)

        if st.session_state.get("editor_current_file") != selected_path:
            st.session_state.editor_current_file = selected_path
            st.session_state.editor_content_cache = selected_file.read_text(encoding="utf-8", errors="ignore")

        editor_widget_key = f"editor_text__{selected_path}"
        edited_text = st.text_area(
            "編輯內容",
            value=st.session_state.get("editor_content_cache", ""),
            height=500,
            key=editor_widget_key,
        )

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            if st.button("💾 儲存修改"):
                backup_path = selected_file.with_suffix(selected_file.suffix + ".bak")
                shutil.copy2(selected_file, backup_path)
                selected_file.write_text(edited_text, encoding="utf-8")
                st.session_state.editor_content_cache = edited_text
                st.success(f"已儲存，備份：{backup_path.name}")

        with c2:
            if st.button("🔄 重新讀取"):
                st.session_state.editor_current_file = selected_path
                st.session_state.editor_content_cache = selected_file.read_text(encoding="utf-8", errors="ignore")
                if editor_widget_key in st.session_state:
                    del st.session_state[editor_widget_key]
                st.rerun()

        with c3:
            if st.button("🧪 語法測試"):
                try:
                    with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False, encoding="utf-8") as tmp:
                        tmp.write(edited_text)
                        tmp_path = tmp.name
                    py_compile.compile(tmp_path, doraise=True)
                    st.success("語法正常 ✓")
                except Exception as e:
                    st.error(str(e))

        with c4:
            confirm_delete = st.checkbox("確認刪除", key="confirm_delete_checkbox")
            if st.button("🗑 刪除"):
                if not confirm_delete:
                    st.error("請先勾選「確認刪除」")
                else:
                    trash_path = selected_file.with_suffix(selected_file.suffix + ".deleted")
                    shutil.move(str(selected_file), str(trash_path))
                    st.success(f"已刪除（保留為：{trash_path.name}）")
                    for k in list(st.session_state.keys()):
                        if k.startswith("editor_text__"):
                            del st.session_state[k]
                    for k in ["editor_current_file", "editor_content_cache"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# 排程設定
# ══════════════════════════════════════════════════
elif page == "排程設定":
    st.markdown('<div class="page-title">排程設定</div><div class="page-subtitle">SCHEDULE CONFIG</div>', unsafe_allow_html=True)

    for task in SCHEDULE_TASKS:
        st.markdown(f'<div class="panel"><div class="panel-title">{task["name"]}｜{task["desc"]}</div>', unsafe_allow_html=True)
        info = load_plist_schedule(task["plist"])
        if info["type"] != "dict":
            st.warning("此 plist 暫不支援直接修改")
            st.markdown("</div>", unsafe_allow_html=True)
            continue

        col1, col2, col3 = st.columns(3)
        with col1:
            day = st.text_input("Day", value=info["day"], key=f"day_{task['label']}")
        with col2:
            hour = st.text_input("Hour", value=info["hour"], key=f"hour_{task['label']}")
        with col3:
            minute = st.text_input("Minute", value=info["minute"], key=f"minute_{task['label']}")

        if st.button(f"儲存 {task['name']}", key=f"save_sched_{task['label']}", use_container_width=True):
            if not hour.isdigit() or not minute.isdigit():
                st.error("Hour / Minute 必須是數字")
            elif day.strip() and not day.isdigit():
                st.error("Day 必須留空或填數字")
            else:
                label, code_out, out, err = save_plist_schedule(task["plist"], day, hour, minute)
                if code_out == 0:
                    st.success(f"已更新 {task['name']}（{label}）")
                else:
                    st.error(f"更新失敗：{err or out}")

        st.markdown(f'<div class="next-run-box">⏭️ 下次執行時間預估：<strong>{calc_next_run(day, hour, minute)}</strong></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════
# 業績報表
# ══════════════════════════════════════════════════
elif page == "業績報表":
    st.markdown('<div class="page-title">業績報表</div><div class="page-subtitle">LATEST + EXECUTION LOG</div>', unsafe_allow_html=True)

    b1, b2, b3 = st.columns(3)

    with b1:
        if st.button("🔄 更新資料", use_container_width=True):
            with st.spinner("更新資料中…"):
                result = generate_sales_report(
                    send_email=False,
                    persist_dashboard=True,
                    trigger="dashboard",
                )
            st.success(f"更新完成：{result['updated_at']}")
            st.rerun()

    with b2:
        if st.button("📧 寄送目前結果", use_container_width=True):
            with st.spinner("寄送中…"):
                result = generate_sales_report(
                    send_email=True,
                    persist_dashboard=True,
                    trigger="dashboard",
                )
            st.success(f"已寄送：{result['updated_at']}")
            st.rerun()

    with b3:
        if st.button("📂 重新讀取已存資料", use_container_width=True):
            st.rerun()

    payload = load_sales_latest_payload()
    df4 = payload["df4"]
    daily_df = payload["daily_df"]
    meta = payload["meta"]
    email_html = payload["email_html"]

    execution_log_df = load_execution_log_for_current_month()
    daily_history_df = load_daily_history_for_current_month()

    updated_at = meta.get("updated_at", "尚未產生資料")
    st.info(f"最新更新時間：{updated_at}")

    total = get_sales_total_row(df4)
    k1, k2, k3, k4 = st.columns(4)
    if total is None:
        k1.metric("本月加總", "0")
        k2.metric("次月加總", "0")
        k3.metric("本月家電加總", "0")
        k4.metric("儲值金", "0")
    else:
        k1.metric("本月加總", sales_fmt_int(total.get("本月加總", 0)))
        k2.metric("次月加總", sales_fmt_int(total.get("次月加總", 0)))
        k3.metric("本月家電加總", sales_fmt_int(total.get("本月家電加總", 0)))
        k4.metric("儲值金", sales_fmt_int(total.get("儲值金", 0)))

    st.markdown('<div class="panel"><div class="panel-title">各區月度摘要</div>', unsafe_allow_html=True)
    if df4.empty:
        st.warning("目前沒有資料")
    else:
        st.dataframe(style_sales_df4(df4), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">當月每日業績總覽</div>', unsafe_allow_html=True)
    if daily_df.empty:
        st.warning("目前沒有資料")
    else:
        st.dataframe(style_sales_daily(daily_df), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">當月累積執行紀錄</div>', unsafe_allow_html=True)
    if execution_log_df.empty:
        st.warning("目前沒有執行紀錄")
    else:
        st.dataframe(style_sales_exec(execution_log_df), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">當月每日業績總覽留存紀錄</div>', unsafe_allow_html=True)
    if daily_history_df.empty:
        st.warning("目前沒有留存紀錄")
    else:
        selectable_ids = daily_history_df["id"].astype(str).tolist()
        default_ids = [x for x in st.session_state.sales_delete_ids if x in selectable_ids]

        selected_ids = st.multiselect(
            "勾選要刪除的紀錄",
            options=selectable_ids,
            default=default_ids,
        )
        st.session_state.sales_delete_ids = selected_ids

        if st.button("🗑 刪除勾選列", use_container_width=True):
            deleted = delete_daily_history_rows(selected_ids)
            st.session_state.sales_delete_ids = []
            st.success(f"已刪除 {deleted} 筆")
            st.rerun()

        display_history = daily_history_df.copy()
        if "daily_json" in display_history.columns:
            display_history = display_history.drop(columns=["daily_json"])

        st.dataframe(style_sales_history(display_history), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">信件預覽</div>', unsafe_allow_html=True)
    if email_html:
        st.components.v1.html(email_html, height=520, scrolling=True)
    else:
        st.info("目前沒有信件內容")
    st.markdown("</div>", unsafe_allow_html=True)

st.caption(f"Lemon Clean Scheduler Console · {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
