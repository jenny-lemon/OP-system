import json
import plistlib
import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from performance_report import (
    generate_sales_report,
    load_execution_log_for_current_month,
    load_daily_history_for_current_month,
    delete_daily_history_rows,
    LATEST_DIR,
)

# =========================================================
# 基本設定
# =========================================================
TZ_TAIPEI = timezone(timedelta(hours=8))

BASE_DIR = Path("/Users/jenny/lemon")
LAUNCH_AGENTS_DIR = Path.home() / "Library/LaunchAgents"
LOG_FILE = BASE_DIR / "cron.log"

# 這 5 個就是主控表要顯示的列
REPORT_TASKS = [
    {
        "name": "排班統計表",
        "key": "schedule_report",
        "script": "schedule_report.py",
        "label": "com.jenny.daily01",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01.plist",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 schedule_report.py',
    },
    {
        "name": "專員班表",
        "key": "staff_schedule",
        "script": "staff_schedule.py",
        "label": "com.jenny.daily01b",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily01b.plist",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 staff_schedule.py',
    },
    {
        "name": "專員個資",
        "key": "staff_info",
        "script": "staff_info.py",
        "label": "com.jenny.daily02b",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02b.plist",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 staff_info.py',
    },
    {
        "name": "訂單資料",
        "key": "orders_report",
        "script": "orders_report.py",
        "label": "com.jenny.daily02",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.daily02.plist",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 orders_report.py',
    },
    {
        "name": "業績報表",
        "key": "performance_report",
        "script": "performance_report.py",
        "label": "com.jenny.sales08",
        "plist": LAUNCH_AGENTS_DIR / "com.jenny.sales08.plist",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 performance_report.py dashboard false',
    },
]

# 其他頁面保留
MANUAL_TASKS = [
    {
        "name": "排班統計表",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 schedule_report.py',
    },
    {
        "name": "專員班表",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 staff_schedule.py',
    },
    {
        "name": "專員個資",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 staff_info.py',
    },
    {
        "name": "訂單資料",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 orders_report.py',
    },
    {
        "name": "業績報表",
        "cmd": f'cd "{BASE_DIR}" && /usr/bin/python3 performance_report.py dashboard false',
    },
]

# =========================================================
# 共用工具
# =========================================================
def now_taipei():
    return datetime.now(TZ_TAIPEI)


def run_shell(cmd: str):
    p = subprocess.run(
        cmd,
        shell=True,
        text=True,
        capture_output=True,
        executable="/bin/bash",
    )
    return p.returncode, p.stdout, p.stderr


def format_dt(dt: Optional[datetime]) -> str:
    if not dt:
        return "—"
    return dt.strftime("%m/%d %H:%M")


def read_last_lines(path: Path, n: int = 80) -> str:
    if not path.exists():
        return "(尚無 log)"
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        return "".join(lines[-n:])
    except Exception as e:
        return f"(讀取失敗) {e}"


def get_launchd_status():
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


def render_status_info(label: str, status_map: dict):
    info = status_map.get(label)
    if not info:
        return {"text": "未載入", "cls": "gray"}

    pid = info["pid"]
    last_exit = info["last_exit"]

    if pid != "-":
        return {"text": f"執行中 PID {pid}", "cls": "yellow"}
    if last_exit == "0":
        return {"text": "正常", "cls": "green"}
    return {"text": f"異常 exit {last_exit}", "cls": "red"}


def load_plist_schedule(plist_path: Path):
    if not plist_path.exists():
        return {"exists": False, "hour": "", "minute": "", "day": ""}

    try:
        with open(plist_path, "rb") as f:
            data = plistlib.load(f)

        interval = data.get("StartCalendarInterval", {})
        if isinstance(interval, list):
            interval = interval[0] if interval else {}

        return {
            "exists": True,
            "hour": str(interval.get("Hour", "")),
            "minute": str(interval.get("Minute", "")),
            "day": str(interval.get("Day", "")),
        }
    except Exception:
        return {"exists": False, "hour": "", "minute": "", "day": ""}


def save_plist_schedule(plist_path: Path, hour: str, minute: str, day: str = ""):
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

    run_shell(f'launchctl bootout gui/$(id -u) "{plist_path}" 2>/dev/null')
    code, out, err = run_shell(f'launchctl bootstrap gui/$(id -u) "{plist_path}"')
    return code, out, err


def calc_next_run(day_str: str, hour_str: str, minute_str: str) -> str:
    try:
        hour = int(hour_str)
        minute = int(minute_str)
        now = now_taipei().replace(tzinfo=None)

        if str(day_str).strip():
            day = int(day_str)
            candidate = now.replace(day=day, hour=hour, minute=minute, second=0, microsecond=0)
            if candidate <= now:
                if now.month == 12:
                    candidate = candidate.replace(year=now.year + 1, month=1)
                else:
                    candidate = candidate.replace(month=now.month + 1)
            return candidate.strftime("%Y-%m-%d %H:%M")

        candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate <= now:
            candidate += timedelta(days=1)
        return candidate.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return "—"


# =========================================================
# 業績報表工具
# =========================================================
def load_sales_latest_payload():
    latest_dir = Path(LATEST_DIR)

    df4_path = latest_dir / "df4.csv"
    daily_path = latest_dir / "daily_df.csv"
    meta_path = latest_dir / "meta.json"
    html_path = latest_dir / "email_preview.html"

    payload = {
        "df4": pd.DataFrame(),
        "daily_df": pd.DataFrame(),
        "meta": {},
        "email_html": "",
    }

    if df4_path.exists():
        payload["df4"] = pd.read_csv(df4_path, encoding="utf-8-sig")
    if daily_path.exists():
        payload["daily_df"] = pd.read_csv(daily_path, encoding="utf-8-sig")
    if meta_path.exists():
        with open(meta_path, "r", encoding="utf-8") as f:
            payload["meta"] = json.load(f)
    if html_path.exists():
        payload["email_html"] = html_path.read_text(encoding="utf-8")

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


# =========================================================
# 畫面樣式
# =========================================================
def inject_styles():
    st.markdown("""
    <style>
    .page-title {
        font-size: 38px;
        font-weight: 900;
        color: #0f172a;
        margin-bottom: 6px;
    }
    .page-subtitle {
        font-size: 14px;
        color: #94a3b8;
        letter-spacing: 0.08em;
        font-family: monospace;
        margin-bottom: 20px;
    }
    .section-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 22px;
        margin-bottom: 18px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.05);
    }
    .section-title {
        font-size: 18px;
        font-weight: 800;
        margin-bottom: 14px;
        color: #0f172a;
    }
    .status-pill {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 700;
    }
    .status-green {
        background: #dcfce7;
        color: #166534;
    }
    .status-yellow {
        background: #fef3c7;
        color: #92400e;
    }
    .status-red {
        background: #fee2e2;
        color: #991b1b;
    }
    .status-gray {
        background: #e5e7eb;
        color: #475569;
    }
    .task-header {
        font-size: 13px;
        color: #94a3b8;
        font-weight: 800;
        padding: 6px 0 10px 0;
        border-bottom: 1px solid #e5e7eb;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)


# =========================================================
# 各頁面
# =========================================================
def render_main_page():
    inject_styles()
    st.markdown('<div class="page-title">主控表</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">SCHEDULE DASHBOARD</div>', unsafe_allow_html=True)

    status_map = get_launchd_status()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">報表任務</div>', unsafe_allow_html=True)

    header = st.columns([2.2, 1.3, 1.6, 1.6, 1.0, 0.9])
    header[0].markdown("**檔案名稱**")
    header[1].markdown("**排程時間**")
    header[2].markdown("**修改排程**")
    header[3].markdown("**執行狀態**")
    header[4].markdown("**下次執行**")
    header[5].markdown("**執行**")

    for task in REPORT_TASKS:
        sched = load_plist_schedule(task["plist"])
        status = render_status_info(task["label"], status_map)

        row = st.columns([2.2, 1.3, 1.6, 1.6, 1.0, 0.9])

        with row[0]:
            st.markdown(f"**{task['name']}**")
            st.caption(task["script"])

        with row[1]:
            if sched["exists"]:
                st.write(f"{sched['hour'].zfill(2)}:{sched['minute'].zfill(2)}")
            else:
                st.write("—")

        with row[2]:
            sub1, sub2, sub3 = st.columns([1, 1, 1])
            hour_key = f"hour_{task['key']}"
            minute_key = f"minute_{task['key']}"
            save_key = f"save_{task['key']}"

            with sub1:
                hour_val = st.text_input(
                    "時",
                    value=sched["hour"],
                    key=hour_key,
                    label_visibility="collapsed",
                    placeholder="時",
                )
            with sub2:
                minute_val = st.text_input(
                    "分",
                    value=sched["minute"],
                    key=minute_key,
                    label_visibility="collapsed",
                    placeholder="分",
                )
            with sub3:
                if st.button("保存", key=save_key, use_container_width=True):
                    if not sched["exists"]:
                        st.error(f"{task['name']} 找不到 plist")
                    elif not hour_val.isdigit() or not minute_val.isdigit():
                        st.error("時間必須是數字")
                    else:
                        code, out, err = save_plist_schedule(
                            task["plist"],
                            hour=hour_val,
                            minute=minute_val,
                            day=sched["day"],
                        )
                        if code == 0:
                            st.success(f"{task['name']} 已更新")
                            st.rerun()
                        else:
                            st.error(err or out or "更新失敗")

        with row[3]:
            cls = status["cls"]
            st.markdown(
                f'<span class="status-pill status-{cls}">{status["text"]}</span>',
                unsafe_allow_html=True,
            )

        with row[4]:
            st.write(calc_next_run(sched["day"], sched["hour"], sched["minute"]))

        with row[5]:
            if st.button("▶", key=f"run_{task['key']}", use_container_width=True):
                with st.spinner(f"執行中：{task['name']}"):
                    code, out, err = run_shell(task["cmd"])
                st.write(f"### {task['name']} 執行結果")
                st.write(f"回傳碼：`{code}`")
                st.text_area(f"stdout_{task['key']}", out or "(無輸出)", height=180)
                st.text_area(f"stderr_{task['key']}", err or "(無錯誤)", height=140)

        st.divider()

    st.markdown('</div>', unsafe_allow_html=True)


def render_sales_page():
    inject_styles()
    st.markdown('<div class="page-title">業績報表</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">LATEST + EXECUTION LOG</div>', unsafe_allow_html=True)

    b1, b2, b3 = st.columns(3)

    with b1:
        if st.button("🔄 更新資料", key="sales_refresh", use_container_width=True):
            with st.spinner("更新資料中…"):
                result = generate_sales_report(
                    send_email=False,
                    persist_dashboard=True,
                    trigger="dashboard",
                )
            st.success(f"更新完成：{result['updated_at']}")
            st.rerun()

    with b2:
        if st.button("📧 寄送目前結果", key="sales_send", use_container_width=True):
            with st.spinner("寄送中…"):
                result = generate_sales_report(
                    send_email=True,
                    persist_dashboard=True,
                    trigger="dashboard",
                )
            st.success(f"已寄送：{result['updated_at']}")
            st.rerun()

    with b3:
        if st.button("📂 重新讀取已存資料", key="sales_reload", use_container_width=True):
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

    st.markdown('<div class="section-card"><div class="section-title">各區月度摘要</div>', unsafe_allow_html=True)
    if df4.empty:
        st.warning("目前沒有資料")
    else:
        st.dataframe(style_sales_df4(df4), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">當月每日業績總覽</div>', unsafe_allow_html=True)
    if daily_df.empty:
        st.warning("目前沒有資料")
    else:
        st.dataframe(style_sales_daily(daily_df), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">當月累積執行紀錄</div>', unsafe_allow_html=True)
    if execution_log_df.empty:
        st.warning("目前沒有執行紀錄")
    else:
        st.dataframe(style_sales_exec(execution_log_df), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">當月每日業績總覽留存紀錄</div>', unsafe_allow_html=True)
    if daily_history_df.empty:
        st.warning("目前沒有留存紀錄")
    else:
        selectable_ids = daily_history_df["id"].astype(str).tolist()
        selected_ids = st.multiselect("勾選要刪除的紀錄", options=selectable_ids)

        if st.button("🗑 刪除勾選列", key="sales_delete_btn", use_container_width=True):
            deleted = delete_daily_history_rows(selected_ids)
            st.success(f"已刪除 {deleted} 筆")
            st.rerun()

        display_history = daily_history_df.copy()
        if "daily_json" in display_history.columns:
            display_history = display_history.drop(columns=["daily_json"])

        st.dataframe(style_sales_history(display_history), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">信件預覽</div>', unsafe_allow_html=True)
    if email_html:
        st.components.v1.html(email_html, height=520, scrolling=True)
    else:
        st.info("目前沒有信件內容")
    st.markdown('</div>', unsafe_allow_html=True)


def render_halfmonth_page():
    inject_styles()
    st.markdown('<div class="page-title">上下半月訂單</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">HALF-MONTH ORDERS</div>', unsafe_allow_html=True)
    st.info("這頁先保留。")


def render_manual_page():
    inject_styles()
    st.markdown('<div class="page-title">手動執行</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">MANUAL TRIGGER</div>', unsafe_allow_html=True)

    selected = st.selectbox("選擇任務", MANUAL_TASKS, format_func=lambda x: x["name"])
    if st.button("▶ 執行", use_container_width=True):
        with st.spinner("執行中…"):
            code, out, err = run_shell(selected["cmd"])
        st.write(f"回傳碼：`{code}`")
        st.text_area("stdout", out or "(無輸出)", height=220)
        st.text_area("stderr", err or "(無錯誤)", height=180)


def render_log_page():
    inject_styles()
    st.markdown('<div class="page-title">Log 監控</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">LOG MONITOR</div>', unsafe_allow_html=True)

    log_choices = {
        "主 log（cron.log）": LOG_FILE,
        "sales08 stderr": BASE_DIR / "launchd_sales08_stderr.log",
        "sales18 stderr": BASE_DIR / "launchd_sales18_stderr.log",
    }

    selected_log = st.selectbox("選擇 log 檔", list(log_choices.keys()))
    raw_log = read_last_lines(log_choices[selected_log], 200)
    st.text_area("內容", raw_log, height=500)


def render_output_page():
    inject_styles()
    st.markdown('<div class="page-title">輸出檔案</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">OUTPUT FILES</div>', unsafe_allow_html=True)

    rows = []
    for name, out_dir in OUTPUT_DIRS.items():
        files = find_latest_files(out_dir, limit=1, category=name)
        latest = files[0] if files else None
        rows.append({
            "分類": name,
            "最新檔案": latest.name if latest else "(無)",
            "時間": file_mtime(latest) if latest else "—",
            "大小": file_size_str(latest) if latest else "—",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_code_page():
    inject_styles()
    st.markdown('<div class="page-title">程式管理</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">CODE MANAGEMENT</div>', unsafe_allow_html=True)

    py_files = list_python_files()
    if not py_files:
        st.warning("找不到任何 .py 檔")
        return

    selected_file = st.selectbox("選擇 Python 檔", py_files, format_func=lambda x: x.name)
    content = selected_file.read_text(encoding="utf-8", errors="ignore")
    st.text_area("內容", content, height=500)


def render_schedule_page():
    inject_styles()
    st.markdown('<div class="page-title">排程設定</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">SCHEDULE CONFIG</div>', unsafe_allow_html=True)

    for task in REPORT_TASKS:
        sched = load_plist_schedule(task["plist"])
        st.markdown(f"### {task['name']}")
        st.write(f"目前時間：{sched['hour'].zfill(2)}:{sched['minute'].zfill(2) if sched['minute'] else '—'}")
        st.write(f"下次執行：{calc_next_run(sched['day'], sched['hour'], sched['minute'])}")
        st.divider()


def render_page(page: str):
    if page == "主控表":
        render_main_page()
    elif page == "業績報表":
        render_sales_page()
    elif page == "上下半月訂單":
        render_halfmonth_page()
    elif page == "手動執行":
        render_manual_page()
    elif page == "Log 監控":
        render_log_page()
    elif page == "輸出檔案":
        render_output_page()
    elif page == "程式管理":
        render_code_page()
    elif page == "排程設定":
        render_schedule_page()
    else:
        render_main_page()
