import os
import re
from datetime import datetime

import requests

from accounts import ACCOUNTS
from paths import PATH_SCHEDULE

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_BASE = "https://backend.lemonclean.com.tw/schedule/export_times"

SAVE_DIR = PATH_SCHEDULE
os.makedirs(SAVE_DIR, exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}


def extract_csrf_token(html: str) -> str:
    patterns = [
        r'<input[^>]*name=["\']_token["\'][^>]*value=["\']([^"\']+)["\']',
        r'<input[^>]*value=["\']([^"\']+)["\'][^>]*name=["\']_token["\']',
    ]
    for pattern in patterns:
        match = re.search(pattern, html, re.IGNORECASE)
        if match:
            return match.group(1)
    raise RuntimeError("找不到 _token，無法登入。")


def login(email: str, password: str) -> requests.Session:
    session = requests.Session()

    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True, timeout=30)
    res.raise_for_status()

    csrf_token = extract_csrf_token(res.text)

    payload = {
        "_token": csrf_token,
        "email": email,
        "password": password,
    }

    login_res = session.post(
        LOGIN_URL,
        data=payload,
        headers=HEADERS,
        allow_redirects=True,
        timeout=30,
    )
    login_res.raise_for_status()

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    return session


def get_month_strings():
    today = datetime.today()
    this_month = today.strftime("%Y-%m")
    today_stamp = today.strftime("%Y%m%d")

    if today.month == 12:
        next_year = today.year + 1
        next_month_num = 1
    else:
        next_year = today.year
        next_month_num = today.month + 1

    next_month = f"{next_year}-{next_month_num:02d}"
    next_month_stamp = f"{next_year}{next_month_num:02d}{today.day:02d}"

    return this_month, next_month, today_stamp, next_month_stamp


def export_schedule(session: requests.Session, month: str, filename: str):
    url = f"{EXPORT_BASE}?month={month}"
    res = session.get(url, headers=HEADERS, allow_redirects=True, timeout=60)

    content_type = res.headers.get("Content-Type", "")
    if res.status_code != 200:
        raise RuntimeError(f"{filename} 下載失敗，status={res.status_code}")

    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{filename} 回傳不是 Excel，Content-Type={content_type}")

    full_path = os.path.join(SAVE_DIR, filename)
    with open(full_path, "wb") as f:
        f.write(res.content)

    print(f"✅ 已下載：{full_path}")


def main():
    this_month, next_month, today_stamp, next_month_stamp = get_month_strings()

    for city in ["台北", "台中"]:
        acc = ACCOUNTS[city]
        print(f"\n=== 處理 {city} ===")

        session = login(acc["email"], acc["password"])
        print("✅ 登入成功")

        current_filename = f"排班統計表{today_stamp}-{city}.xlsx"
        next_filename = f"排班統計表{next_month_stamp}-{city}.xlsx"

        export_schedule(session, this_month, current_filename)
        export_schedule(session, next_month, next_filename)


if __name__ == "__main__":
    main()
