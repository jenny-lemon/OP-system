"""
dashboard_main.py
─────────────────
所有頁面的 render 函式與輔助函式。
不含 st.set_page_config、不含 CSS、不含 Navigation。
由 opapp.py 負責 layout + CSS，再呼叫 render_page()。
"""

import json
import plistlib
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, Callable, Dict, List

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

# ── Runtime detection ─────────────────────────────────────────────────────────

TZ_TAIPEI = timezone(timedelta(hours=8))
PROJECT_DIR = Path(__file__).resolve().parent
LOCAL_BASE_DIR = Path("/Users/jenny/lemon")
BASE_DIR = LOCAL_BASE_DIR if LOCAL_BASE_DIR.exists() else PROJECT_DIR
IS_LOCAL = LOCAL_BASE_DIR.exists()

LOG_FILE = BASE_DIR / "cron.log"
LAUNCH_AGENTS_DIR = Path.home() / "Library/LaunchAgents"
CONFIG_FILE = BASE_DIR / "dashboard_config.json"
PYTHON_CMD = sys.executable or "python3"

# ── Task definitions ──────────────────────────────────────────────────────────

OUTPUT_DIRS = {
    "排班統計表": Path(PATH_SCHEDULE),
    "專員班表":   Path(PATH_CLEANER_SCHEDULE),
    "專員個資":   Path(PATH_CLEANER_DATA),
    "訂單資料":   Path(PATH_ORDER),
    "業績報表":   Path(PATH_REPORT),
    "預收":       Path(PATH_VIP),
    "儲值金結算": Path(PATH_VIP),
    "儲值金預收": Path(PATH_VIP),
    "上下半月訂單": Path(PATH_HR),
    "已退款":     Path(PATH_HR),
}

MAIN_REPORT_TASKS = [
    {
        "name": "排班統計表",
        "task_key": "schedule_report",
        "script": "schedule_report.py",
        "label": "com.jenny.daily01",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01.plist",
        "default_hour": "01", "default_minute": "10",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" schedule_report.py',
    },
    {
        "name": "專員班表",
        "task_key": "staff_schedule",
        "script": "staff_schedule.py",
        "label": "com.jenny.daily01b",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01b.plist",
        "default_hour": "01", "default_minute": "20",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_schedule.py',
    },
    {
        "name": "專員個資",
        "task_key": "staff_info",
        "script": "staff_info.py",
        "label": "com.jenny.daily02b",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02b.plist",
        "default_hour": "01", "default_minute": "30",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_info.py',
    },
    {
        "name": "訂單資料",
        "task_key": "orders_report",
        "script": "orders_report.py",
        "label": "com.jenny.daily02",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02.plist",
        "default_hour": "01", "default_minute": "40",
        "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" orders_report.py',
    },
    {
        "name": "業績報表",
        "task_key": "performance_report",
        "script": "performance_report.py",
        "label": "com.jenny.sales08",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.sales08.plist",
        "default_hour": "08", "default_minute": "00",
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

# ── Shared helpers ────────────────────────────────────────────────────────────

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
    if s < 1024:         return f"{s} B"
    if s < 1024 * 1024:  return f"{s/1024:.1f} KB"
    return f"{s/1024/1024:.1f} MB"


def find_latest_files(base_dir: Path, limit: int = 10) -> List[Path]:
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
            return str(v).zfill(2) if str(v).strip() != "" else fb
        return {
            "exists": True,
            "hour":   _z(iv.get("Hour",   default_hour),   default_hour),
            "minute": _z(iv.get("Minute", default_minute), default_minute),
            "day":    str(iv.get("Day", "")),
            "source": "plist",
        }
    except Exception:
        return fallback


def save_plist_schedule(plist_path: Path, hour: str, minute: str, day: str = ""):
    with open(plist_path, "rb") as f:
        data = plistlib.load(f)
    iv: dict = {}
    if day.strip():
        iv["Day"] = int(day)
    iv["Hour"]   = int(hour)
    iv["Minute"] = int(minute)
    data["StartCalendarInterval"] = iv
    with open(plist_path, "wb") as f:
        plistlib.dump(data, f)
    run_shell(f'launchctl bootout gui/$(id -u) "{plist_path}" 2>/dev/null')
    return run_shell(f'launchctl bootstrap gui/$(id -u) "{plist_path}"')


def calc_next_run(day_str: str, hour_str: str, minute_str: str) -> str:
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


def load_sales_latest_payload() -> dict:
    latest_dir = Path(LATEST_DIR)
    payload: dict = {
        "df4": pd.DataFrame(),
        "daily_df": pd.DataFrame(),
        "meta": {},
        "email_html": "",
    }
    for key, fname in [("df4", "df4.csv"), ("daily_df", "daily_df.csv")]:
        p = latest_dir / fname
        if p.exists():
            try:
                payload[key] = pd.read_csv(p, encoding="utf-8-sig")
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


def _fmt_int(x) -> str:
    try:    return f"{int(float(x)):,}"
    except: return "—"


def _fmt_pct(x) -> str:
    try:    return f"{float(x):.2%}"
    except: return "—"


# ── HTML table renderer (guaranteed right-align) ──────────────────────────────

def render_html_table(
    df: pd.DataFrame,
    right_cols: set,
    pct_cols: set,
    int_cols: set,
) -> str:
    """
    Convert a DataFrame to an HTML table.
    - right_cols: columns to right-align
    - pct_cols:   columns formatted as percentage
    - int_cols:   columns formatted as integer with comma separator
    """

    def _cell(val, col):
        if pd.isna(val) or str(val).strip() in ("", "nan"):
            return "—"
        if col in pct_cols:
            return _fmt_pct(val)
        if col in int_cols:
            return _fmt_int(val)
        return str(val)

    TH = (
        "padding:10px 14px;"
        "font-size:10.5px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;"
        "color:#64748b;border-bottom:2px solid #e2e8f0;white-space:nowrap;background:#fafafa;"
    )
    TD_BASE = "padding:9px 14px;font-size:13px;color:#1e293b;border-bottom:1px solid #f1f5f9;white-space:nowrap;"
    TD_R = TD_BASE + "text-align:right;font-variant-numeric:tabular-nums;font-family:'DM Mono','Menlo',monospace;"
    TD_L = TD_BASE + "text-align:left;"

    ths = "".join(
        f'<th style="{TH}text-align:{"right" if c in right_cols else "left"}">{c}</th>'
        for c in df.columns
    )
    rows = []
    for _, row in df.iterrows():
        tds = "".join(
            f'<td style="{TD_R if c in right_cols else TD_L}">{_cell(row[c], c)}</td>'
            for c in df.columns
        )
        rows.append(f"<tr>{tds}</tr>")

    return (
        '<div style="overflow-x:auto;border:1.5px solid #e2e8f0;border-radius:12px;">'
        f'<table style="width:100%;border-collapse:collapse;background:#fff;">'
        f'<thead><tr>{ths}</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody>'
        "</table></div>"
    )


def _badge(label: str, cls: str) -> str:
    CLS = {"green": "b-green", "yellow": "b-yellow", "red": "b-red", "gray": "b-gray"}
    return f'<span class="badge {CLS.get(cls,"b-gray")}">{label}</span>'


# ── Page: 主控表 ──────────────────────────────────────────────────────────────

def render_main_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">排程主控表</div>'
        '<div class="page-subtitle">Schedule Dashboard</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    if "task_results" not in st.session_state:
        st.session_state.task_results = {}

    status_map = get_launchd_status()

    count_ok  = sum(1 for t in MAIN_REPORT_TASKS if launchd_badge(t["label"], status_map)[1] == "green")
    count_run = sum(1 for t in MAIN_REPORT_TASKS if launchd_badge(t["label"], status_map)[1] == "yellow")
    count_err = len(MAIN_REPORT_TASKS) - count_ok - count_run
    ran_today = sum(1 for v in st.session_state.task_results.values() if v)

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi-card blue">
        <div class="kpi-label">Total</div><div class="kpi-value">{len(MAIN_REPORT_TASKS)}</div>
        <div class="kpi-sub">已設定排程</div>
      </div>
      <div class="kpi-card green">
        <div class="kpi-label">Normal</div><div class="kpi-value">{count_ok}</div>
        <div class="kpi-sub">上次退出正常</div>
      </div>
      <div class="kpi-card amber">
        <div class="kpi-label">Running</div><div class="kpi-value">{count_run}</div>
        <div class="kpi-sub">目前執行中</div>
      </div>
      <div class="kpi-card {"red" if count_err > 0 else "blue"}">
        <div class="kpi-label">今日已執行</div><div class="kpi-value">{ran_today}</div>
        <div class="kpi-sub">本次 session</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Task table ────────────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 報表任務</div>', unsafe_allow_html=True)

    # Column header
    st.markdown("""
    <div style="display:grid;grid-template-columns:1.5fr .8fr 1.5fr 1fr 1fr 1.1fr .55fr;
                gap:0;padding:0 4px 10px;border-bottom:1.5px solid #e2e8f0;margin-bottom:4px;">
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">任務 / 腳本</span>
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">launchd</span>
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">排程時間（改）</span>
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">目前設定</span>
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">上次結果</span>
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;">下次執行</span>
      <span style="font-size:9.5px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;text-align:center;">執行</span>
    </div>
    """, unsafe_allow_html=True)

    for task in MAIN_REPORT_TASKS:
        sched = load_plist_schedule(task["plist"], task["default_hour"], task["default_minute"])
        ld_text, ld_cls = launchd_badge(task["label"], status_map)
        result = st.session_state.task_results.get(task["task_key"])

        if result is None:
            res_badge = _badge("尚未執行", "gray")
        elif result["code"] == 0:
            res_badge = _badge("✓ 成功", "green")
        else:
            res_badge = _badge("✗ 失敗", "red")

        c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, .8, 1.5, 1, 1, 1.1, .55])

        with c1:
            st.markdown(
                f"<span style='font-weight:700;color:#0f172a;font-size:13px'>{task['name']}</span><br>"
                f"<span style='font-size:11px;color:#94a3b8;font-family:monospace'>{task['script']}</span>",
                unsafe_allow_html=True,
            )

        with c2:
            st.markdown(_badge(ld_text, ld_cls), unsafe_allow_html=True)

        with c3:
            ci1, ci2, ci3 = st.columns([1, 1, .7])
            with ci1:
                h_val = st.text_input("時", value=sched["hour"],
                                      key=f'h_{task["task_key"]}', label_visibility="collapsed", placeholder="HH")
            with ci2:
                m_val = st.text_input("分", value=sched["minute"],
                                      key=f'm_{task["task_key"]}', label_visibility="collapsed", placeholder="MM")
            with ci3:
                st.markdown('<div class="save-btn">', unsafe_allow_html=True)
                if st.button("💾", key=f'save_{task["task_key"]}', use_container_width=True):
                    if not IS_LOCAL:
                        st.warning("雲端環境無法修改本機 plist")
                    elif not task["plist"].exists():
                        st.error(f"找不到 plist：{task['plist'].name}")
                    elif not h_val.isdigit() or not m_val.isdigit():
                        st.error("時間必須是數字")
                    else:
                        code, _, err = save_plist_schedule(task["plist"], h_val, m_val, sched["day"])
                        if code == 0:
                            st.success("✓ 已更新排程")
                            st.rerun()
                        else:
                            st.error(err or "更新失敗")
                st.markdown("</div>", unsafe_allow_html=True)

        with c4:
            note = "" if sched["source"] == "plist" else \
                '<span style="font-size:10px;color:#f59e0b"> (預設)</span>'
            st.markdown(
                f'<span style="font-family:monospace;font-size:13.5px;font-weight:600;color:#1e293b">'
                f'{sched["hour"]}:{sched["minute"]}</span>{note}',
                unsafe_allow_html=True,
            )

        with c5:
            st.markdown(res_badge, unsafe_allow_html=True)
            if result:
                st.markdown(
                    f'<span style="font-size:10.5px;color:#94a3b8">{result.get("ran_at","")[-8:]}</span>',
                    unsafe_allow_html=True,
                )

        with c6:
            st.markdown(
                f'<span class="next-run">🕐 {calc_next_run(sched["day"], sched["hour"], sched["minute"])}</span>',
                unsafe_allow_html=True,
            )

        with c7:
            st.markdown('<div class="run-btn">', unsafe_allow_html=True)
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

        st.markdown("<div style='height:2px'></div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)  # close section-card

    # ── Execution result panels ───────────────────────────────────────────────
    results_to_show = [
        (task, st.session_state.task_results.get(task["task_key"]))
        for task in MAIN_REPORT_TASKS
        if st.session_state.task_results.get(task["task_key"])
    ]

    if results_to_show:
        st.markdown(
            '<div style="font-size:12px;font-weight:700;color:#64748b;'
            'letter-spacing:.08em;text-transform:uppercase;margin:22px 0 10px">'
            '▼ 執行結果</div>',
            unsafe_allow_html=True,
        )
        for task, r in results_to_show:
            rc = r["code"]
            rc_badge = _badge(f"exit {rc}", "green" if rc == 0 else "red")
            st.markdown(
                f'<div class="exec-panel">'
                f'<div class="exec-panel-title">'
                f'▶&nbsp;{r["name"]}'
                f'&emsp;<span style="font-size:12px;color:#94a3b8;font-weight:500">{r["ran_at"]}</span>'
                f'&emsp;{rc_badge}'
                f'</div>',
                unsafe_allow_html=True,
            )
            # Show stderr immediately on failure
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
                if not r["stdout"].strip() and not r["stderr"].strip():
                    st.markdown('<div class="log-box"><span class="log-normal">(無輸出)</span></div>', unsafe_allow_html=True)

            st.markdown("</div>", unsafe_allow_html=True)


# ── Page: 業績報表 ────────────────────────────────────────────────────────────

def render_sales_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">業績報表</div>'
        '<div class="page-subtitle">Latest Data · Send Later</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    result = None
    c1, c2, c3 = st.columns([1, 1, 1.5])
    with c1:
        update_btn = st.button("🔄 更新資料", use_container_width=True)
    with c2:
        send_btn = st.button("📧 寄送目前結果", use_container_width=True)
    with c3:
        if st.button("📂 重新讀取已存資料", use_container_width=True):
            st.rerun()

    if update_btn:
        with st.spinner("更新資料中…"):
            result = generate_sales_report(send_email=False, persist_dashboard=True, trigger="dashboard")

    # Load data
    if result is not None:
        df4         = result.get("df4", pd.DataFrame())
        daily_df    = result.get("daily_df", pd.DataFrame())
        email_html  = result.get("email_html", "")
        updated_at  = result.get("updated_at", "")
        exec_log_df = result.get("execution_log_df", pd.DataFrame())
        error_msg   = result.get("error")
    else:
        payload     = load_sales_latest_payload()
        df4         = payload.get("df4", pd.DataFrame())
        daily_df    = payload.get("daily_df", pd.DataFrame())
        meta        = payload.get("meta", {})
        email_html  = payload.get("email_html", "")
        updated_at  = meta.get("updated_at", "尚未產生資料") if isinstance(meta, dict) else "—"
        exec_log_df = load_execution_log_for_current_month()
        error_msg   = meta.get("error") if isinstance(meta, dict) else None
        if payload.get("df4_error"):
            st.warning(f"df4.csv 讀取錯誤：{payload['df4_error']}")
        if payload.get("daily_df_error"):
            st.warning(f"daily_df.csv 讀取錯誤：{payload['daily_df_error']}")

    # Send email
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

    # KPI
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

    # ── 各區月度摘要 ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📊 各區月度摘要</div>', unsafe_allow_html=True)

    if df4.empty:
        st.markdown('<div class="empty-state"><span class="icon">📭</span>目前沒有資料，請先按「更新資料」</div>', unsafe_allow_html=True)
    else:
        INT4   = {"本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"}
        PCT4   = {"本月佔比", "次月佔比"}
        RIGHT4 = INT4 | PCT4
        st.markdown(render_html_table(df4, right_cols=RIGHT4, pct_cols=PCT4, int_cols=INT4), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── 當月每日業績總覽 ──────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📅 當月每日業績總覽</div>', unsafe_allow_html=True)

    # Debug info
    daily_csv = Path(LATEST_DIR) / "daily_df.csv"
    parts = []
    if daily_csv.exists():
        parts.append(f"daily_df.csv 存在（{file_size_str(daily_csv)}，更新：{file_mtime(daily_csv)}）")
    else:
        parts.append("⚠️ daily_df.csv 不存在")
    parts.append(f"載入：{len(daily_df)} 行 × {len(daily_df.columns)} 欄")
    if not daily_df.empty and len(daily_df.columns):
        parts.append(f"欄位：{', '.join(daily_df.columns.tolist()[:6])}{'…' if len(daily_df.columns) > 6 else ''}")
    st.caption("  ·  ".join(parts))

    if daily_df.empty:
        reason = (
            "daily_df.csv 不存在，請先按「更新資料」產生資料。"
            if not daily_csv.exists()
            else "CSV 存在但無資料列。可能原因：日期欄解析失敗，或抓取資料中無有效日期。請查看 Log 確認。"
        )
        st.markdown(
            f'<div class="empty-state"><span class="icon">📭</span>{reason}</div>',
            unsafe_allow_html=True,
        )
    else:
        DAILY_COLS = [
            "日期",
            "台北業績", "台北佔比", "台中業績", "台中佔比",
            "桃園業績", "桃園佔比", "新竹業績", "新竹佔比",
            "高雄業績", "高雄佔比", "全區合計",
        ]
        exist_cols  = [c for c in DAILY_COLS if c in daily_df.columns]
        df_show     = daily_df[exist_cols].copy()
        INT_D   = {c for c in exist_cols if "業績" in c or c == "全區合計"}
        PCT_D   = {c for c in exist_cols if "佔比" in c}
        RIGHT_D = INT_D | PCT_D
        st.markdown(render_html_table(df_show, right_cols=RIGHT_D, pct_cols=PCT_D, int_cols=INT_D), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── 當月累積執行紀錄 ──────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📝 當月累積執行紀錄</div>', unsafe_allow_html=True)

    if exec_log_df.empty:
        st.markdown('<div class="empty-state"><span class="icon">📋</span>目前沒有執行紀錄</div>', unsafe_allow_html=True)
    else:
        exec_ids = exec_log_df["id"].astype(str).tolist()
        sel_ids  = st.multiselect("勾選要刪除的執行紀錄", options=exec_ids, key="del_exec")
        if st.button("🗑 刪除勾選列", key="del_exec_btn", use_container_width=True):
            deleted = delete_execution_log_rows(sel_ids)
            st.success(f"已刪除 {deleted} 筆")
            st.rerun()
        INT_E = {"本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"}
        st.markdown(render_html_table(exec_log_df, right_cols=INT_E, pct_cols=set(), int_cols=INT_E), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── 信件預覽 ──────────────────────────────────────────────────────────────
    if email_html:
        with st.expander("📧 信件預覽"):
            st.components.v1.html(email_html, height=520, scrolling=True)


# ── Page: 上下半月訂單 ────────────────────────────────────────────────────────

def render_halfmonth_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">上下半月訂單</div>'
        '<div class="page-subtitle">Half-Month Orders</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.info("這頁先保留。")


# ── Page: 手動執行 ────────────────────────────────────────────────────────────

def render_manual_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">手動執行</div>'
        '<div class="page-subtitle">Manual Trigger</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">▶ 選擇任務執行</div>', unsafe_allow_html=True)

    selected = st.selectbox("選擇任務", MANUAL_TASKS, format_func=lambda x: x["name"])
    if st.button("▶ 執行", use_container_width=True):
        with st.spinner("執行中…"):
            rc, out, err = run_shell(selected["cmd"])
        rc_b = _badge(f"exit {rc}", "green" if rc == 0 else "red")
        st.markdown(f"回傳碼：{rc_b}", unsafe_allow_html=True)
        if out.strip():
            st.markdown('<div class="exec-label">STDOUT</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="log-box">{highlight_log(out)}</div>', unsafe_allow_html=True)
        if err.strip():
            st.markdown('<div class="exec-label">STDERR</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="log-box">{highlight_log(err)}</div>', unsafe_allow_html=True)
        if not out.strip() and not err.strip():
            st.markdown('<div class="log-box"><span class="log-normal">(無輸出)</span></div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ── Page: Log 監控 ────────────────────────────────────────────────────────────

def render_log_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">Log 監控</div>'
        '<div class="page-subtitle">Log Monitor</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    log_choices = {
        "主 log（cron.log）":     LOG_FILE,
        "sales08 stderr":         BASE_DIR / "launchd_sales08_stderr.log",
        "sales18 stderr":         BASE_DIR / "launchd_sales18_stderr.log",
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
    raw_log  = read_last_lines(log_path, n_lines)
    st.markdown(f'<div class="log-meta">📄 {log_path}&emsp;·&emsp;更新：{file_mtime(log_path)}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="log-box">{highlight_log(raw_log)}</div>', unsafe_allow_html=True)


# ── Page: 輸出檔案 ────────────────────────────────────────────────────────────

def render_output_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">輸出檔案監控</div>'
        '<div class="page-subtitle">Output Files</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    rows = []
    for name, out_dir in OUTPUT_DIRS.items():
        files  = find_latest_files(out_dir, limit=1)
        latest = files[0] if files else None
        rows.append({
            "分類":     name,
            "最新檔案": latest.name if latest else "(無)",
            "時間":     file_mtime(latest),
            "大小":     file_size_str(latest),
            "資料夾":   str(out_dir),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ── Page: 程式管理 ────────────────────────────────────────────────────────────

def render_code_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">程式管理</div>'
        '<div class="page-subtitle">Code Management</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.info("這頁先保留。")


# ── Page: 排程設定 ────────────────────────────────────────────────────────────

def render_schedule_page():
    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">排程設定</div>'
        '<div class="page-subtitle">Schedule Config</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.info(f"Python：{PYTHON_CMD}")

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
            minute = st.text_input("Minute", value=sched["minute"], key=f"smin_{task['label']}")
        with col4:
            st.markdown(
                f'<div class="next-run" style="margin-top:28px">'
                f'🕐 下次執行：{calc_next_run(sched["day"], hour, minute)}</div>',
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
                code, _, err = save_plist_schedule(task["plist"], hour, minute, day)
                st.success(f"✓ 已更新 {task['name']}") if code == 0 else st.error(err or "更新失敗")

        st.markdown("</div>", unsafe_allow_html=True)


# ── Router ────────────────────────────────────────────────────────────────────

def render_page(page: str):
    dispatch = {
        "主控表":       render_main_page,
        "業績報表":     render_sales_page,
        "上下半月訂單": render_halfmonth_page,
        "手動執行":     render_manual_page,
        "Log 監控":     render_log_page,
        "輸出檔案":     render_output_page,
        "程式管理":     render_code_page,
        "排程設定":     render_schedule_page,
    }
    dispatch.get(page, render_main_page)()

    # Footer
    st.markdown(
        f'<div class="footer-cap">'
        f'Lemon Clean Scheduler Console &nbsp;·&nbsp; {now_taipei().strftime("%Y-%m-%d %H:%M:%S")}'
        f'</div>',
        unsafe_allow_html=True,
    )
