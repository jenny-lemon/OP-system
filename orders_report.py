import os
import calendar
import tempfile
from datetime import datetime

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

# Google Drive
GDRIVE_FOLDER_ID = "1QnOJzn-xmZ_oAMoiM6Qnfk3Y2CWuM1c4"
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

    login_res = session.post(
        LOGIN_URL,
        data=payload,
        headers=HEADERS,
        allow_redirects=True,
    )

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


def get_drive_service():
    try:
        creds_dict = dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])

        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=GDRIVE_SCOPES,
        )

        return build("drive", "v3", credentials=creds)

    except Exception as e:
        raise RuntimeError(f"Google Drive 初始化失敗：{e}")


def upload_to_gdrive(local_path):
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


def export_order(session, city, export_url, file_date):
    res = session.get(export_url, headers=HEADERS, allow_redirects=True)

    if res.status_code != 200:
        raise RuntimeError(f"{city} 匯出失敗，status={res.status_code}")

    content_type = res.headers.get("Content-Type", "")
    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{city} 不是 Excel，Content-Type={content_type}")

    filename = f"{file_date}訂單-{city}.xls"

    with tempfile.TemporaryDirectory() as tmpdir:
        full_path = os.path.join(tmpdir, filename)

        with open(full_path, "wb") as f:
            f.write(res.content)

        print(f"✅ 已下載到暫存：{full_path}")
        print(f"📦 檔案大小：{os.path.getsize(full_path)} bytes")

        upload_to_gdrive(full_path)


def main():
    city, email, password = get_selected_region_account()
    session = requests.Session()

    start, end, file_date = get_date_range()
    export_url = build_export_url(start, end)

    print(f"📌 服務日期：{start} ~ {end}")
    print(f"\n=== 處理 {city} ===")

    try:
        login(session, email, password)
        export_order(session, city, export_url, file_date)
        print(f"✅ {city} 完成")

    except Exception as e:
        print(f"❌ {city} 失敗：{e}")
        raise


if __name__ == "__main__":
    main()
