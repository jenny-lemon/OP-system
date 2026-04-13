import json
import os
import subprocess
from datetime import datetime, timezone, timedelta
from pathlib import Path

TZ = timezone(timedelta(hours=8))

CONFIG = {
    "daily": [
        {"id": "d1", "label": "排班統計表", "script": "schedule_report.py", "args": [], "schedule": "01:10"},
        {"id": "d2", "label": "專員班表", "script": "staff_schedule.py", "args": [], "schedule": "01:20"},
        {"id": "d3", "label": "專員個資", "script": "staff_info.py", "args": [], "schedule": "01:30"},
        {"id": "d4", "label": "當月次月訂單", "script": "orders_report.py", "args": [], "schedule": "01:40"},
        {"id": "d5", "label": "業績報表", "script": "performance_report.py", "args": [], "schedule": "08:00"},
    ],
    "monthly": [
        {"id": "m1", "label": "上半月訂單", "script": "half_month_orders.py", "args": ["1"], "schedule": "每月15日18:15"},
        {"id": "m2", "label": "下半月訂單", "script": "half_month_orders.py", "args": ["2"], "schedule": "每月底18:15"},
        {"id": "m3", "label": "已退款", "script": "refund_report.py", "args": [], "schedule": "月底18:30"},
        {"id": "m4", "label": "預收", "script": "prepaid_report.py", "args": [], "schedule": "月初00:10"},
        {"id": "m5", "label": "儲值金結算", "script": "stored_value_settlement.py", "args": [], "schedule": "月初00:20"},
        {"id": "m6", "label": "儲值金預收", "script": "stored_value_prepaid.py", "args": [], "schedule": "月初00:30"},
    ],
}


def is_last_day(dt):
    return (dt + timedelta(days=1)).month != dt.month


def match_jobs(now):
    matched = []
    hhmm = now.strftime("%H:%M")
    day = now.day

    for job in CONFIG["daily"]:
        if job["schedule"] == hhmm:
            matched.append(job)

    for job in CONFIG["monthly"]:
        s = job["schedule"]

        if s == "每月15日18:15" and day == 15 and hhmm == "18:15":
            matched.append(job)

        if s == "每月底18:15" and is_last_day(now) and hhmm == "18:15":
            matched.append(job)

        if s == "月底18:30" and is_last_day(now) and hhmm == "18:30":
            matched.append(job)

        if s == "月初00:10" and day == 1 and hhmm == "00:10":
            matched.append(job)

        if s == "月初00:20" and day == 1 and hhmm == "00:20":
            matched.append(job)

        if s == "月初00:30" and day == 1 and hhmm == "00:30":
            matched.append(job)

    return matched


def main():
    now = datetime.now(TZ).replace(second=0, microsecond=0)
    print(f"台北時間：{now:%Y-%m-%d %H:%M}")

    jobs = match_jobs(now)

    if not jobs:
        print("沒有任務")
        return

    base_dir = Path(__file__).resolve().parent

    for job in jobs:
        print(f"開始執行：{job['label']}")
        subprocess.run(["python", job["script"], *job["args"]], check=True, cwd=base_dir)


if __name__ == "__main__":
    main()
