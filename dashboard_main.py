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

TZ_TAIPEI = timezone(timedelta(hours=8))

PROJECT_DIR = Path(__file__).resolve().parent
LOCAL_BASE_DIR = Path("/Users/jenny/lemon")

if LOCAL_BASE_DIR.exists():
    BASE_DIR = LOCAL_BASE_DIR
    IS_LOCAL_RUNTIME = True
else:
    BASE_DIR = PROJECT_DIR
    IS_LOCAL_RUNTIME = False

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
    "專員個資": Path(PATH_CLEANER_DATA),
    "訂單資料": Path(PATH_ORDER),
    "業績報表": Path(PATH_REPORT),
    "預收": Path(PATH_VIP),
    "儲值金結算": Path(PATH_VIP),
    "儲值金預收": Path(PATH_VIP),
    "上下半月訂單": Path(PATH_HR),
    "已退款": Path(PATH_HR),
}

CATEGORY_MATCHERS: Dict[str, Callable[[Path], bool]] = {
    "排班統計表": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "專員班表": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "專員個資": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "訂單資料": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".gsheet"},
    "業績報表": lambda p: p.suffix.lower() in {".xlsx", ".xls", ".csv", ".html", ".json"},
    "預收": lambda p: "預收" in p.name and "儲值金預收" not in p.name,
    "儲值金結算": lambda p: "儲值金結算" in p.name,
    "儲值金預收": lambda p: "儲值金預收" in p.name,
    "上下半月訂單": lambda p: "訂單-" in p.name and "已退款" not in p.name,
    "已退款": lambda p: "已退款" in p.name,
}

PYTHON_CMD = sys.executable or "python3"

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
    {"name": "專員班表", "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_schedule.py'},
    {"name": "專員個資", "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" staff_info.py'},
    {"name": "訂單資料", "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" orders_report.py'},
    {"name": "業績報表", "cmd": f'cd "{BASE_DIR}" && "{PYTHON_CMD}" performance_report.py dashboard false'},
]


def now_taipei():
    return datetime.now(TZ_TAIPEI)


def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def run_shell(cmd: str):
    p = subprocess.run(
        cmd,
        shell=True,
        text=True,
        capture_output=True,
        executable="/bin/bash",
    )
    return p.returncode, p.stdout, p.stderr


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
    ts = datetime.fromtimestamp(path.stat().st_mtime, tz=TZ_TAIPEI)
    return ts.strftime("%m/%d %H:%M")


def file_size_str(path: Optional[Path]) -> str:
    if not path or not path.exists():
        return "-"
    size = path.stat().st_size
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size/1024:.1f} KB"
    return f"{size/1024/1024:.1f} MB"


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


def render_launchd_status(label: str, status_map: dict):
    info = status_map.get(label)
    if not info:
        return ("未載入", "gray")
    pid = info["pid"]
    last_exit = info["last_exit"]
    if pid != "-":
        return (f"執行中 PID {pid}", "yellow")
    if last_exit == "0":
        return ("正常", "green")
    return (f"異常 exit {last_exit}", "red")


def load_plist_schedule(plist_path: Path, default_hour: str = "", default_minute: str = ""):
    fallback = {
        "exists": False,
        "hour": default_hour,
        "minute": default_minute,
        "day": "",
        "source": "default",
    }

    if not plist_path.exists():
        return fallback

    try:
        with open(plist_path, "rb") as f:
            data = plistlib.load(f)

        interval = data.get("StartCalendarInterval", {})
        if isinstance(interval, list):
            interval = interval[0] if interval else {}

        return {
            "exists": True,
            "hour": str(interval.get("Hour", default_hour)).zfill(2) if str(interval.get("Hour", default_hour)) != "" else "",
            "minute": str(interval.get("Minute", default_minute)).zfill(2) if str(interval.get("Minute", default_minute)) != "" else "",
            "day": str(interval.get("Day", "")),
            "source": "plist",
        }
    except Exception:
        return fallback


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


def get_sales_total_row(df4):
    if df4.empty:
        return None
    total = df4[df4["城市"] == "加總"]
    if total.empty:
        return None
    return total.iloc[0]


def inject_styles():
    st.markdown("""
    <style>
    .page-title {
        font-size: 42px;
        font-weight: 900;
        color: #0f172a;
        margin-bottom: 6px;
    }
    .page-subtitle {
        font-size: 16px;
        color: #94a3b8;
        letter-spacing: 0.08em;
        font-family: monospace;
        margin-bottom: 18px;
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
        font-size: 22px;
        font-weight: 900;
        margin-bottom: 16px;
        color: #0f172a;
    }
    .status-pill {
        display: inline-block;
        padding: 7px 14px;
        border-radius: 999px;
        font-size: 14px;
        font-weight: 800;
    }
    .status-green { background: #dcfce7; color: #166534; }
    .status-yellow { background: #fef3c7; color: #92400e; }
    .status-red { background: #fee2e2; color: #991b1b; }
    .status-gray { background: #e5e7eb; color: #475569; }
    </style>
    """, unsafe_allow_html=True)


def render_main_page():
    inject_styles()
    st.markdown('<div class="page-title">主控表</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">SCHEDULE DASHBOARD</div>', unsafe_allow_html=True)

    status_map = get_launchd_status()

    if "task_results" not in st.session_state:
        st.session_state.task_results = {}

    st.markdown('<div class="section-card"><div class="section-title">報表任務</div>', unsafe_allow_html=True)

    header = st.columns([2.0, 1.1, 1.8, 1.2, 1.1, 1.3, 0.8])
    header[0].markdown("**檔案名稱**")
    header[1].markdown("**目前排程**")
    header[2].markdown("**修改排程**")
    header[3].markdown("**執行狀態**")
    header[4].markdown("**執行結果**")
    header[5].markdown("**下次執行**")
    header[6].markdown("**執行**")

    for task in MAIN_REPORT_TASKS:
        sched = load_plist_schedule(task["plist"], task["default_hour"], task["default_minute"])
        status_text, status_cls = render_launchd_status(task["label"], status_map)
        result_data = st.session_state.task_results.get(task["task_key"])

        row = st.columns([2.0, 1.1, 1.8, 1.2, 1.1, 1.3, 0.8])

        with row[0]:
            st.markdown(f"**{task['name']}**")
            st.caption(task["script"])

        with row[1]:
            st.write(f'{sched["hour"]}:{sched["minute"]}' if sched["hour"] and sched["minute"] else "—")
            if sched.get("source") == "default":
                st.caption("預設值")

        with row[2]:
            c1, c2, c3 = st.columns([1, 1, 0.8])
            with c1:
                hour_val = st.text_input("時", value=sched["hour"], key=f'hour_{task["task_key"]}', label_visibility="collapsed")
            with c2:
                minute_val = st.text_input("分", value=sched["minute"], key=f'minute_{task["task_key"]}', label_visibility="collapsed")
            with c3:
                if st.button("💾", key=f'save_{task["task_key"]}', use_container_width=True):
                    if not task["plist"].exists():
                        st.warning(f'{task["name"]}：目前雲端環境無法直接修改本機 plist')
                    elif not hour_val.isdigit() or not minute_val.isdigit():
                        st.error("時間必須是數字")
                    else:
                        code, out, err = save_plist_schedule(task["plist"], hour=hour_val, minute=minute_val, day=sched["day"])
                        if code == 0:
                            st.success(f'{task["name"]} 排程已更新')
                            st.rerun()
                        else:
                            st.error(err or out or "更新失敗")

        with row[3]:
            st.markdown(f'<span class="status-pill status-{status_cls}">{status_text}</span>', unsafe_allow_html=True)

        with row[4]:
            if result_data is None:
                st.write("—")
            elif result_data["code"] == 0:
                st.success("成功")
            else:
                st.error("失敗")

        with row[5]:
            st.write(calc_next_run(sched["day"], sched["hour"], sched["minute"]))

        with row[6]:
            if st.button("▶", key=f'run_{task["task_key"]}', use_container_width=True):
                with st.spinner(f'執行中：{task["name"]}'):
                    code, out, err = run_shell(task["cmd"])
                st.session_state.task_results[task["task_key"]] = {
                    "name": task["name"],
                    "code": code,
                    "stdout": out,
                    "stderr": err,
                    "ran_at": now_taipei().strftime("%Y-%m-%d %H:%M:%S"),
                }
                st.rerun()

        result_data = st.session_state.task_results.get(task["task_key"])
        if result_data is not None:
            st.markdown(f"**{task['name']} 執行結果**")
            st.write(f'執行時間：{result_data["ran_at"]}')
            st.write(f'回傳碼：{result_data["code"]}')
            if result_data["code"] == 0:
                st.success("執行成功")
            else:
                st.error("執行失敗")
                fail_reason = result_data["stderr"] or result_data["stdout"] or "沒有收到錯誤訊息"
                st.markdown("**失敗原因**")
                st.code(fail_reason)

            with st.expander("查看完整 stdout / stderr"):
                st.text_area(f'stdout_{task["task_key"]}', result_data["stdout"] or "(無輸出)", height=140)
                st.text_area(f'stderr_{task["task_key"]}', result_data["stderr"] or "(無錯誤)", height=140)

        st.divider()

    st.markdown('</div>', unsafe_allow_html=True)


def render_sales_page():
    inject_styles()
    st.markdown('<div class="page-title">業績報表</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">LATEST + EXECUTION LOG</div>', unsafe_allow_html=True)

    result = None

    b1, b2, b3 = st.columns(3)

    with b1:
        if st.button("🔄 更新資料", key="sales_refresh", use_container_width=True):
            with st.spinner("更新資料中…"):
                result = generate_sales_report(
                    send_email=False,
                    persist_dashboard=True,
                    trigger="dashboard",
                )

    with b2:
        if st.button("📧 寄送目前結果", key="sales_send", use_container_width=True):
            with st.spinner("寄送中…"):
                result = generate_sales_report(
                    send_email=True,
                    persist_dashboard=True,
                    trigger="dashboard",
                )

    with b3:
        if st.button("📂 重新讀取已存資料", key="sales_reload", use_container_width=True):
            st.rerun()

    if result is not None:
        df4 = result.get("df4", pd.DataFrame())
        daily_df = result.get("daily_df", pd.DataFrame())
        email_html = result.get("email_html", "")
        updated_at = result.get("updated_at", now_taipei().strftime("%Y-%m-%d %H:%M:%S"))
        execution_log_df = result.get("execution_log_df", pd.DataFrame())
        error_msg = result.get("error")
    else:
        payload = load_sales_latest_payload()
        df4 = payload.get("df4", pd.DataFrame())
        daily_df = payload.get("daily_df", pd.DataFrame())
        meta = payload.get("meta", {})
        email_html = payload.get("email_html", "")
        updated_at = meta.get("updated_at", "尚未產生資料")
        execution_log_df = load_execution_log_for_current_month()
        error_msg = meta.get("error") if isinstance(meta, dict) else None

    if error_msg:
        st.error(error_msg)

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

    # 各區月度摘要
    st.markdown('<div class="section-card"><div class="section-title">各區月度摘要</div>', unsafe_allow_html=True)
    if df4.empty:
        st.warning("目前沒有資料")
    else:
        summary_style = (
            df4.style
            .format({
                "本月加總": "{:,.0f}",
                "次月加總": "{:,.0f}",
                "本月家電加總": "{:,.0f}",
                "次月家電加總": "{:,.0f}",
                "儲值金": "{:,.0f}",
                "本月佔比": "{:.2%}",
                "次月佔比": "{:.2%}",
            })
            .set_properties(subset=["城市"], **{"text-align": "left"})
            .set_properties(
                subset=[
                    "本月加總", "本月佔比", "次月加總", "次月佔比",
                    "本月家電加總", "次月家電加總", "儲值金"
                ],
                **{"text-align": "right"}
            )
        )
        st.dataframe(summary_style, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 當月每日業績總覽
    st.markdown('<div class="section-card"><div class="section-title">當月每日業績總覽</div>', unsafe_allow_html=True)
    if daily_df.empty:
        st.warning("目前沒有資料")
    else:
        cols = [
            "日期",
            "台北業績", "台北佔比",
            "台中業績", "台中佔比",
            "桃園業績", "桃園佔比",
            "新竹業績", "新竹佔比",
            "高雄業績", "高雄佔比",
            "全區合計",
        ]
        exist_cols = [c for c in cols if c in daily_df.columns]
        df_show = daily_df[exist_cols].copy()

        daily_style = (
            df_show.style
            .format({
                "台北業績": "{:,.0f}",
                "台中業績": "{:,.0f}",
                "桃園業績": "{:,.0f}",
                "新竹業績": "{:,.0f}",
                "高雄業績": "{:,.0f}",
                "全區合計": "{:,.0f}",
                "台北佔比": "{:.2%}",
                "台中佔比": "{:.2%}",
                "桃園佔比": "{:.2%}",
                "新竹佔比": "{:.2%}",
                "高雄佔比": "{:.2%}",
            })
            .set_properties(subset=["日期"], **{"text-align": "left"})
            .set_properties(subset=[c for c in exist_cols if c != "日期"], **{"text-align": "right"})
        )
        st.dataframe(daily_style, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 當月累積執行紀錄（這裡加刪除功能）
    st.markdown('<div class="section-card"><div class="section-title">當月累積執行紀錄</div>', unsafe_allow_html=True)
    if execution_log_df.empty:
        st.warning("目前沒有執行紀錄")
    else:
        exec_ids = execution_log_df["id"].astype(str).tolist()
        selected_exec_ids = st.multiselect("勾選要刪除的執行紀錄", options=exec_ids)

        if st.button("🗑 刪除勾選列", key="delete_execution_log_btn", use_container_width=True):
            from performance_report import delete_execution_log_rows
            deleted = delete_execution_log_rows(selected_exec_ids)
            st.success(f"已刪除 {deleted} 筆執行紀錄")
            st.rerun()

        exec_style = (
            execution_log_df.style
            .format({
                "本月加總": "{:,.0f}",
                "次月加總": "{:,.0f}",
                "本月家電加總": "{:,.0f}",
                "次月家電加總": "{:,.0f}",
                "儲值金": "{:,.0f}",
            })
            .set_properties(subset=["id", "執行時間", "來源"], **{"text-align": "left"})
            .set_properties(
                subset=["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"],
                **{"text-align": "right"}
            )
        )
        st.dataframe(exec_style, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 信件預覽
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
    st.info("這頁先保留。")


def render_schedule_page():
    inject_styles()
    st.markdown('<div class="page-title">排程設定</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">SCHEDULE CONFIG</div>', unsafe_allow_html=True)

    for task in MAIN_REPORT_TASKS:
        sched = load_plist_schedule(task["plist"], task["default_hour"], task["default_minute"])
        st.markdown(f"### {task['name']}")
        st.write(f"目前時間：{sched['hour']}:{sched['minute']}")
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
