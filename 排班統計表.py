"""
opapp.py  ──  營運報表控制台 v3
Features:
  • 固定頂部導覽列（捲動後仍固定）
  • 地區切換，帳密從 accounts.py 讀取（本機，不上傳雲端）
  • 排程主控表：每日 / 月排程執行狀態一覽
  • 手動執行：單項 / 批次，帶入地區帳密
  • 輸出報表：掃描 output/ 目錄，顯示檔案狀態
  • 腳本管理：線上新增 / 修改 / 刪除 / 編輯腳本內容

帳密安全做法：
  accounts.py 加入 .gitignore → 永遠不上傳雲端
  opapp.py 只做 from accounts import ACCOUNTS，可以安全上傳
"""

import os
import tempfile
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from accounts import ACCOUNTS

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


def get_drive_service():
    if not os.path.exists(GOOGLE_SERVICE_ACCOUNT_FILE):
        raise RuntimeError(
            f"找不到 service account 檔案：{GOOGLE_SERVICE_ACCOUNT_FILE}"
        )

    creds = service_account.Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_FILE,
        scopes=GDRIVE_SCOPES,
    )
    return build("drive", "v3", credentials=creds)


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
    this_month, next_month, today_stamp, next_month_stamp = get_month_strings()

    for city in ["台北", "台中"]:
        acc = ACCOUNTS[city]
        print(f"\n=== 處理 {city} ===")

        try:
            session = login(acc["email"], acc["password"])
            print("✅ 登入成功")

            current_filename = f"排班統計表{today_stamp}-{city}.xlsx"
            next_filename = f"排班統計表{next_month_stamp}-{city}.xlsx"

            export_schedule(session, this_month, current_filename)
            export_schedule(session, next_month, next_filename)

        except Exception as e:
            print(f"❌ {city} 失敗：{e}")
            raise


if __name__ == "__main__":
    main()
