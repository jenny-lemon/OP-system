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

# 👉 Google Drive（訂單資料資料夾）
GDRIVE_FOLDER_ID = "1QnOJzn-xmZ_oAMoiM6Qnfk3Y2CWuM1c4"
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

# 👉 台北時區
TZ = timezone(timedelta(hours=8))


# =========================
# 帳密
# =========================
def load_accounts():
    return [
        (
            "台北",
            st.secrets["accounts"]["taipei"]["email"],
            st.secrets["accounts"]["taipei"]["password"],
        ),
        (
            "台中",
            st.secrets["accounts"]["taichung"]["email"],
            st.secrets["accounts"]["taichung"]["password"],
        ),
    ]


# =========================
# 登入
# =========================
def login(email: str, password: str) -> requests.Session:
    session = requests.Session()

    res = session.get(LOGIN_URL, headers=HEADERS)
    soup = BeautifulSoup(res.text, "html.parser")

    token = soup.find("input", {"name": "_token"}).get("value")

    payload = {
        "_token": token,
        "email": email,
        "password": password,
    }

    login_res = session.post(LOGIN_URL, data=payload, headers=HEADERS)

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    return session


# =========================
# 日期區間
# =========================
def get_date_range():
    now = datetime.now(TZ)

    start = f"{now.year}-{now.month:02d}-01"

    if now.month == 12:
        ny, nm = now.year + 1, 1
    else:
        ny, nm = now.year, now.month + 1

    last_day = calendar.monthrange(ny, nm)[1]
    end = f"{ny}-{nm:02d}-{last_day:02d}"

    file_date = now.strftime("%Y%m%d")

    return start, end, file_date


# =========================
# 建 URL
# =========================
def build_export_url(start, end):
    params = {
        "clean_date_s": start,
        "clean_date_e": end,
        "purchase_status": "1",
    }

    return requests.Request("GET", EXPORT_URL, params=params).prepare().url


# =========================
# Google Drive
# =========================
def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
        scopes=GDRIVE_SCOPES,
    )
    return build("drive", "v3", credentials=creds)


def upload_to_gdrive(local_path: str):
    service = get_drive_service()
    filename = os.path.basename(local_path)

    file_metadata = {
        "name": filename,
        "parents": [GDRIVE_FOLDER_ID],
    }

    media = MediaFileUpload(local_path, resumable=True)

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()

    print(f"☁️ 已上傳：{created['name']}")
    return created["id"]


# =========================
# 匯出訂單
# =========================
def export_order(session, city, export_url, file_date):
    res = session.get(export_url, headers=HEADERS)

    if res.status_code != 200:
        raise RuntimeError(f"{city} 匯出失敗，status={res.status_code}")

    content_type = res.headers.get("Content-Type", "")
    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{city} 不是 Excel，Content-Type={content_type}")

    filename = f"{file_date}訂單-{city}.xls"

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, filename)

        with open(path, "wb") as f:
            f.write(res.content)

        print(f"✅ 已下載：{path}")
        print(f"📦 檔案大小：{os.path.getsize(path)} bytes")

        upload_to_gdrive(path)


# =========================
# 主程式
# =========================
def main():
    print("🔥 訂單資料匯出")

    start, end, file_date = get_date_range()
    export_url = build_export_url(start, end)

    print(f"📌 服務日期：{start} ~ {end}")

    regions = load_accounts()

    for city, email, password in regions:
        print(f"\n=== {city} ===")

        try:
            session = login(email, password)
            print("✅ 登入成功")

            export_order(session, city, export_url, file_date)

            print(f"✅ {city} 完成")

        except Exception as e:
            print(f"❌ {city} 失敗：{e}")
            raise


if __name__ == "__main__":
    main()
