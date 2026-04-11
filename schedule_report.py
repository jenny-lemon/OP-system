import os
import tempfile
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_BASE = "https://backend.lemonclean.com.tw/schedule/export_times"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

# Google Drive
GDRIVE_FOLDER_ID = "1cNWX1eL6SzkjJH8qQQyoz6cUoFHoVrRq"
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv(
    "GOOGLE_SERVICE_ACCOUNT_FILE",
    "google_service_account.json",
)
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_selected_region_account():
    region_name = os.getenv("REGION_NAME", "").strip()
    region_email = os.getenv("REGION_EMAIL", "").strip()
    region_password = os.getenv("REGION_PASSWORD", "").strip()

    if not region_name:
        raise RuntimeError("未指定地區，請先在介面選擇台北或台中")

    if not region_email or not region_password:
        raise RuntimeError(f"{region_name} 找不到帳密，請檢查 Streamlit secrets 或地區設定")

    return region_name, region_email, region_password


def login(email: str, password: str) -> requests.Session:
    session = requests.Session()

    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    soup = BeautifulSoup(res.text, "html.parser")

    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise RuntimeError("找不到 _token，無法登入。")

    csrf_token = token_input.get("value")

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
    )

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


import json
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

def get_drive_service():
    try:
        creds_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]

        credentials = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive"]
        )

        return build("drive", "v3", credentials=credentials)

    except Exception as e:
        raise RuntimeError(f"Google Drive 初始化失敗: {e}")


def upload_to_gdrive(local_path: str, folder_id: str):
    service = get_drive_service()
    filename = os.path.basename(local_path)

    file_metadata = {
        "name": filename,
        "parents": [folder_id],
    }

    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name",
        supportsAllDrives=True,
    ).execute()

    print(f"☁️ 已上傳到 Google Drive：{created['name']} (file_id={created['id']})")
    return created["id"]


def export_schedule(session: requests.Session, month: str, filename: str):
    url = f"{EXPORT_BASE}?month={month}"
    res = session.get(url, headers=HEADERS, allow_redirects=True)

    content_type = res.headers.get("Content-Type", "")
    if res.status_code != 200:
        raise RuntimeError(f"{filename} 下載失敗，status={res.status_code}")

    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{filename} 回傳不是 Excel，Content-Type={content_type}")

    with tempfile.TemporaryDirectory() as tmpdir:
        full_path = os.path.join(tmpdir, filename)

        with open(full_path, "wb") as f:
            f.write(res.content)

        print(f"✅ 已下載到暫存：{full_path}")
        print(f"📦 檔案大小：{os.path.getsize(full_path)} bytes")

        upload_to_gdrive(full_path, GDRIVE_FOLDER_ID)


def main():
    city, email, password = get_selected_region_account()
    this_month, next_month, today_stamp, next_month_stamp = get_month_strings()

    print(f"\n=== 處理 {city} ===")

    try:
        session = login(email, password)
        print("✅ 登入成功")

        current_filename = f"排班統計表{today_stamp}-{city}.xlsx"
        next_filename = f"排班統計表{next_month_stamp}-{city}.xlsx"

        export_schedule(session, this_month, current_filename)
        export_schedule(session, next_month, next_filename)

        print(f"✅ {city} 全部完成")

    except Exception as e:
        print(f"❌ {city} 失敗：{e}")
        raise


if __name__ == "__main__":
    main()
