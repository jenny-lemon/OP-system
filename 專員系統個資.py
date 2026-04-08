import os
from datetime import datetime
import requests
from bs4 import BeautifulSoup
from accounts import ACCOUNTS

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_URL = "https://backend.lemonclean.com.tw/cleaner1/export"

SAVE_DIR = "/Users/jenny/Library/CloudStorage/GoogleDrive-jenny@lemonclean.com.tw/我的雲端硬碟/lemon_jenny/Jenny@lemon程式/專員系統個資"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}


def login(session: requests.Session, email: str, password: str):
    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    soup = BeautifulSoup(res.text, "html.parser")

    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise RuntimeError("找不到 _token，無法登入")

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


def export_cleaner_data(session: requests.Session, city: str):
    save_dir = SAVE_DIR
    os.makedirs(save_dir, exist_ok=True)

    today_str = datetime.today().strftime("%Y%m%d")
    filename = f"{today_str}專員系統個資-{city}.xlsx"
    full_path = os.path.join(save_dir, filename)

    res = session.get(EXPORT_URL, headers=HEADERS, allow_redirects=True)

    content_type = res.headers.get("Content-Type", "")
    if res.status_code != 200:
        raise RuntimeError(f"{city} 匯出失敗，status={res.status_code}")

    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{city} 回傳不是 Excel，Content-Type={content_type}")

    with open(full_path, "wb") as f:
        f.write(res.content)

    print(f"✅ 已下載：{full_path}")


def main():
    for city in ["台北", "台中"]:
        acc = ACCOUNTS[city]
        session = requests.Session()

        try:
            print(f"\n=== 處理 {city} ===")
            login(session, acc["email"], acc["password"])
            export_cleaner_data(session, city)
        except Exception as e:
            print(f"❌ {city} 失敗：{e}")


if __name__ == "__main__":
    main()
