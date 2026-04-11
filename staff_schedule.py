import os
from datetime import datetime
import requests
from bs4 import BeautifulSoup

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_BASE = "https://backend.lemonclean.com.tw/cleaner1/export_all"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

# 👉 改這：直接寫或之後拉 config
API_LIMIT = 10000


def get_account():
    region = os.getenv("REGION_NAME")
    email = os.getenv("REGION_EMAIL")
    password = os.getenv("REGION_PASSWORD")

    if not region:
        raise RuntimeError("未指定地區")

    return region, email, password


def login(session, email, password):
    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    soup = BeautifulSoup(res.text, "html.parser")

    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise RuntimeError("找不到 _token")

    token = token_input.get("value")

    payload = {
        "_token": token,
        "email": email,
        "password": password,
    }

    login_res = session.post(LOGIN_URL, data=payload, headers=HEADERS, allow_redirects=True)

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    print(f"✅ 登入成功：{email}")


def get_months():
    today = datetime.today()

    this_month = today.strftime("%Y-%m")
    this_file_date = today.strftime("%Y%m%d")

    if today.month == 12:
        next_year = today.year + 1
        next_month_num = 1
    else:
        next_year = today.year
        next_month_num = today.month + 1

    next_month = f"{next_year}-{next_month_num:02d}"
    next_file_date = f"{next_year}{next_month_num:02d}{today.day:02d}"

    return this_month, next_month, this_file_date, next_file_date


def export_cleaner_schedule(session, month, city, filename):
    save_dir = "output/cleaner_schedule"
    os.makedirs(save_dir, exist_ok=True)

    url = f"{EXPORT_BASE}?month={month}&limit={API_LIMIT}"
    res = session.get(url, headers=HEADERS, allow_redirects=True)

    content_type = res.headers.get("Content-Type", "")
    if res.status_code != 200:
        raise RuntimeError(f"{city} 匯出失敗，status={res.status_code}")

    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{city} 回傳不是 Excel，Content-Type={content_type}")

    full_path = os.path.join(save_dir, filename)

    with open(full_path, "wb") as f:
        f.write(res.content)

    print(f"✅ 已下載：{full_path}")


def main():
    city, email, password = get_account()

    this_month, next_month, this_date, next_date = get_months()

    session = requests.Session()

    print(f"\n=== {city} ===")

    try:
        login(session, email, password)

        current_filename = f"{this_date}專員班表-{city}.xls"
        next_filename = f"{next_date}專員班表-{city}.xls"

        export_cleaner_schedule(session, this_month, city, current_filename)
        export_cleaner_schedule(session, next_month, city, next_filename)

        print(f"✅ {city} 完成")

    except Exception as e:
        print(f"❌ {city} 失敗：{e}")
        raise


if __name__ == "__main__":
    main()
