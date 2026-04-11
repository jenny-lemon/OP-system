import os
import tempfile
from datetime import datetime

import requests
import streamlit as st
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

LOGIN_URL = "https://backend.lemonclean.com.tw/login"

# 🔥 👉 這裡改成你的業績報表 API
EXPORT_URL = "EXPORT_URL = "https://backend.lemonclean.com.tw/purchase"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

# 👉 業績報表資料夾
GDRIVE_FOLDER_ID = "1_b2EjuCAZ6qdlzUjY_PiecbpV7MKBC2t"
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


# =========================
# 帳密（從 opapp 傳）
# =========================
def get_account():
    region = os.getenv("REGION_NAME")
    email = os.getenv("REGION_EMAIL")
    password = os.getenv("REGION_PASSWORD")

    if not region:
        raise RuntimeError("未指定地區")

    if not email or not password:
        raise RuntimeError(f"{region} 找不到帳密")

    return region, email, password


# =========================
# 登入
# =========================
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

    login_res = session.post(
        LOGIN_URL,
        data=payload,
        headers=HEADERS,
        allow_redirects=True,
    )

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    print(f"✅ 登入成功：{email}")


# =========================
# Google Drive
# =========================
def get_drive_service():
    creds_dict = dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])

    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=GDRIVE_SCOPES,
    )

    return build("drive", "v3", credentials=creds)


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


# =========================
# 匯出業績
# =========================
def export_performance(session, city):
    today_str = datetime.today().strftime("%Y%m%d")
    filename = f"{today_str}業績報表-{city}.xlsx"

    res = session.get(EXPORT_URL, headers=HEADERS, allow_redirects=True)

    if res.status_code != 200:
        raise RuntimeError(f"{city} 匯出失敗，status={res.status_code}")

    content_type = res.headers.get("Content-Type", "")
    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{city} 不是 Excel，Content-Type={content_type}")

    with tempfile.TemporaryDirectory() as tmpdir:
        full_path = os.path.join(tmpdir, filename)

        with open(full_path, "wb") as f:
            f.write(res.content)

        print(f"✅ 已下載到暫存：{full_path}")
        print(f"📦 檔案大小：{os.path.getsize(full_path)} bytes")

        upload_to_gdrive(full_path)


# =========================
# 主程式
# =========================
def main():
    city, email, password = get_account()
    session = requests.Session()

    print(f"\n=== 處理 {city} ===")

    try:
        login(session, email, password)
        export_performance(session, city)
        print(f"✅ {city} 完成")

    except Exception as e:
        print(f"❌ {city} 失敗：{e}")
        raise


if __name__ == "__main__":
    main()
