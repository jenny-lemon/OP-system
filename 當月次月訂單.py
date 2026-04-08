import os
import calendar
from datetime import datetime
import requests
from bs4 import BeautifulSoup
from accounts import ACCOUNTS

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_URL = "https://backend.lemonclean.com.tw/purchase/export_order"

SAVE_DIR = "/Users/jenny/Library/CloudStorage/GoogleDrive-jenny@lemonclean.com.tw/我的雲端硬碟/lemon_jenny/Jenny@lemon程式/訂單資料"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}


def login(session: requests.Session, email: str, password: str):
    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    soup = BeautifulSoup(res.text, "html.parser")

    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise RuntimeError("找不到 _token")

    csrf_token = token_input.get("value")

    payload = {
        "_token": csrf_token,
        "email": email,
        "password": password,
    }

    login_res = session.post(LOGIN_URL, data=payload, headers=HEADERS, allow_redirects=True)

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    print(f"✅ 登入成功：{email}")


def get_date_range():
    today = datetime.today()

    start = f"{today.year}-{today.month:02d}-01"

    if today.month == 12:
        next_year = today.year + 1
        next_month = 1
    else:
        next_year = today.year
        next_month = today.month + 1

    last_day = calendar.monthrange(next_year, next_month)[1]
    end = f"{next_year}-{next_month:02d}-{last_day:02d}"

    file_date = today.strftime("%Y%m%d")

    return start, end, file_date


def build_export_url(start, end):
    params = {
        "clean_date_s": start,
        "clean_date_e": end,
        "purchase_status": "1",
    }

    req = requests.Request("GET", EXPORT_URL, params=params).prepare()
    return req.url


def export_order(session, city, export_url, file_date):
    save_dir = SAVE_DIR
    os.makedirs(save_dir, exist_ok=True)

    res = session.get(export_url, headers=HEADERS, allow_redirects=True)

    if res.status_code != 200:
        raise RuntimeError(f"{city} 匯出失敗，status={res.status_code}")

    content_type = res.headers.get("Content-Type", "")
    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{city} 不是 Excel，Content-Type={content_type}")

    filename = f"{file_date}訂單-{city}.xls"
    path = os.path.join(save_dir, filename)

    with open(path, "wb") as f:
        f.write(res.content)

    print(f"✅ 已下載：{path}")


def main():
    start, end, file_date = get_date_range()
    export_url = build_export_url(start, end)

    print(f"📌 服務日期：{start} ~ {end}")

    for city in ["台北", "台中"]:
        acc = ACCOUNTS[city]
        session = requests.Session()

        try:
            print(f"\n=== {city} ===")
            login(session, acc["email"], acc["password"])
            export_order(session, city, export_url, file_date)
        except Exception as e:
            print(f"❌ {city} 失敗：{e}")


if __name__ == "__main__":
    main()
