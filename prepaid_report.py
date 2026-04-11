import os
import calendar
import tempfile
from datetime import datetime, timezone, timedelta

import requests
import streamlit as st
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_URL = "https://backend.lemonclean.com.tw/purchase/export_order"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

TZ = timezone(timedelta(hours=8))

# 👉 預收主資料夾
GDRIVE_FOLDER_ID = "1VCb_y-zBA7tm9SF1s7GeixZVweWteWIc"
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


# =========================
# 帳密
# =========================
def load_accounts():
    return [
        ("台北", st.secrets["accounts"]["taipei"]["email"], st.secrets["accounts"]["taipei"]["password"]),
        ("台中", st.secrets["accounts"]["taichung"]["email"], st.secrets["accounts"]["taichung"]["password"]),
    ]


# =========================
# 登入
# =========================
def login(email, password):
    session = requests.Session()

    res = session.get(LOGIN_URL, headers=HEADERS)
    soup = BeautifulSoup(res.text, "html.parser")
    token = soup.find("input", {"name": "_token"}).get("value")

    payload = {
        "_token": token,
        "email": email,
        "password": password,
    }

    res = session.post(LOGIN_URL, data=payload, headers=HEADERS)

    if "login" in res.url.lower():
        raise RuntimeError("登入失敗")

    return session


# =========================
# 日期邏輯（🔥核心）
# =========================
def get_date_ranges():
    now = datetime.now(TZ)
    y, m = now.year, now.month

    # 👉 上月付款
    if m == 1:
        py, pm = y - 1, 12
    else:
        py, pm = y, m - 1

    paid_at_s = f"{py}-{pm:02d}-01"
    paid_at_e = f"{py}-{pm:02d}-{calendar.monthrange(py, pm)[1]:02d}"

    # 👉 本月服務
    clean_date_s = f"{y}-{m:02d}-01"

    # 👉 +4個月
    ey, em = y, m + 4
    while em > 12:
        em -= 12
        ey += 1

    clean_date_e = f"{ey}-{em:02d}-{calendar.monthrange(ey, em)[1]:02d}"

    file_tag = f"{py}{pm:02d}"

    return {
        "paid_at_s": paid_at_s,
        "paid_at_e": paid_at_e,
        "clean_date_s": clean_date_s,
        "clean_date_e": clean_date_e,
        "file_tag": file_tag,
    }


# =========================
# URL
# =========================
def build_export_url(rng, keyword=""):
    params = {
        "keyword": keyword,
        "paid_at_s": rng["paid_at_s"],
        "paid_at_e": rng["paid_at_e"],
        "clean_date_s": rng["clean_date_s"],
        "clean_date_e": rng["clean_date_e"],
        "purchase_status": "1",
        "p_board": "on",
    }

    req = requests.Request("GET", EXPORT_URL, params=params).prepare()
    return req.url


# =========================
# Drive
# =========================
def get_drive():
    creds = service_account.Credentials.from_service_account_info(
        dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
        scopes=GDRIVE_SCOPES,
    )
    return build("drive", "v3", credentials=creds)


def upload_file(path):
    drive = get_drive()

    media = MediaFileUpload(path, resumable=True)

    drive.files().create(
        body={
            "name": os.path.basename(path),
            "parents": [GDRIVE_FOLDER_ID],
        },
        media_body=media,
        supportsAllDrives=True,
    ).execute()

    print("☁️ 上傳完成")


# =========================
# 主程式
# =========================
def main():
    print("🔥 預收報表")

    rng = get_date_ranges()

    print(f"付款日期：{rng['paid_at_s']} ~ {rng['paid_at_e']}")
    print(f"服務日期：{rng['clean_date_s']} ~ {rng['clean_date_e']}")

    accounts = load_accounts()

    for city, email, password in accounts:
        print(f"\n=== {city} ===")

        try:
            session = login(email, password)
            print("✅ 登入成功")

            keyword = "新竹" if city == "新竹" else ""
            url = build_export_url(rng, keyword)

            res = session.get(url, headers=HEADERS)
            res.raise_for_status()

            filename = f"{rng['file_tag']}預收-{city}.xlsx"

            with tempfile.TemporaryDirectory() as tmp:
                path = os.path.join(tmp, filename)

                with open(path, "wb") as f:
                    f.write(res.content)

                print(f"✅ 已下載：{path}")
                upload_file(path)

        except Exception as e:
            print(f"❌ {city} 失敗：{e}")
            raise


if __name__ == "__main__":
    main()
