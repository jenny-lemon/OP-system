import os
import shutil
import tempfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

try:
    import streamlit as st
    HAS_STREAMLIT = True
except Exception:
    HAS_STREAMLIT = False


LOGIN_URL = "https://backend.lemonclean.com.tw/login"
EXPORT_BASE = "https://backend.lemonclean.com.tw/schedule/export_times"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

# Google Drive
GDRIVE_FOLDER_ID = "1V0IjoJqHlnkGb3Oq70Cil63pQ9j8r2Xv"
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

# 台北時區
TZ = timezone(timedelta(hours=8))

# 是否在 GitHub Actions
IS_GITHUB = os.getenv("GITHUB_ACTIONS") == "true"

# 本機輸出資料夾（只有本機跑時才會用）
LOCAL_OUTPUT_DIR = Path(
    "/Users/jenny/Library/CloudStorage/GoogleDrive-jenny@lemonclean.com.tw/.shortcut-targets-by-id/1zbu45AG1adMzz24HPdi_tLfh2Tncw_Br/排班統計表"
)


def log(msg: str):
    now_str = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now_str}] {msg}")


def get_secret(path_list, env_name=None, required=True, default=None):
    """
    依序嘗試：
    1. streamlit secrets
    2. 環境變數
    """
    # 先讀 streamlit secrets
    if HAS_STREAMLIT:
        try:
            cur = st.secrets
            for key in path_list:
                cur = cur[key]
            if cur is not None and str(cur) != "":
                return cur
        except Exception:
            pass

    # 再讀環境變數
    if env_name:
        value = os.getenv(env_name)
        if value not in (None, ""):
            return value

    if required:
        raise RuntimeError(f"讀不到設定值：{'/'.join(path_list)}")
    return default


def load_accounts():
    return [
        (
            "台北",
            get_secret(["accounts", "taipei", "email"], env_name="TAIPEI_EMAIL"),
            get_secret(["accounts", "taipei", "password"], env_name="TAIPEI_PASSWORD"),
        ),
        (
            "台中",
            get_secret(["accounts", "taichung", "email"], env_name="TAICHUNG_EMAIL"),
            get_secret(["accounts", "taichung", "password"], env_name="TAICHUNG_PASSWORD"),
        ),
    ]


def get_service_account_info():
    """
    優先用 streamlit secrets 的 GOOGLE_SERVICE_ACCOUNT
    否則用環境變數 GOOGLE_SERVICE_ACCOUNT_JSON
    """
    if HAS_STREAMLIT:
        try:
            creds_dict = dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
            if creds_dict:
                return creds_dict
        except Exception:
            pass

    json_str = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if json_str:
        import json
        return json.loads(json_str)

    raise RuntimeError("找不到 GOOGLE_SERVICE_ACCOUNT 設定")


def login(email: str, password: str) -> requests.Session:
    session = requests.Session()

    log(f"開始登入：{email}")

    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True, timeout=60)
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
        timeout=60,
    )

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    log(f"登入成功：{email}")
    return session


def get_month_strings():
    now = datetime.now(TZ)

    this_month = now.strftime("%Y-%m")
    today_stamp = now.strftime("%Y%m%d")

    if now.month == 12:
        next_year = now.year + 1
        next_month_num = 1
    else:
        next_year = now.year
        next_month_num = now.month + 1

    next_month = f"{next_year}-{next_month_num:02d}"
    next_month_stamp = f"{next_year}{next_month_num:02d}{now.day:02d}"

    return this_month, next_month, today_stamp, next_month_stamp


def get_drive_service():
    try:
        creds_dict = get_service_account_info()

        credentials = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=GDRIVE_SCOPES,
        )

        service = build("drive", "v3", credentials=credentials)
        log("Google Drive 初始化成功")
        return service

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

    log(f"準備上傳到 Google Drive：{filename}")

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name",
        supportsAllDrives=True,
    ).execute()

    log(f"☁️ 已上傳到 Google Drive：{created['name']} (file_id={created['id']})")
    return created["id"]


def save_to_local_if_possible(temp_file_path: str, filename: str):
    """
    只有在本機跑時，才另外複製到本機資料夾
    GitHub Actions 直接跳過
    """
    if IS_GITHUB:
        log("目前在 GitHub Actions，略過本機存檔")
        return

    try:
        LOCAL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        local_file = LOCAL_OUTPUT_DIR / filename
        shutil.copy2(temp_file_path, local_file)
        log(f"💾 已另存本機：{local_file}")
    except Exception as e:
        raise RuntimeError(f"本機存檔失敗：{e}")


def export_schedule(session: requests.Session, month: str, filename: str):
    url = f"{EXPORT_BASE}?month={month}"
    log(f"開始下載：month={month} / filename={filename}")

    res = session.get(url, headers=HEADERS, allow_redirects=True, timeout=120)

    content_type = res.headers.get("Content-Type", "")
    if res.status_code != 200:
        raise RuntimeError(f"{filename} 下載失敗，status={res.status_code}")

    if "excel" not in content_type.lower() and "octet-stream" not in content_type.lower():
        raise RuntimeError(f"{filename} 回傳不是 Excel，Content-Type={content_type}")

    with tempfile.TemporaryDirectory() as tmpdir:
        full_path = os.path.join(tmpdir, filename)

        with open(full_path, "wb") as f:
            f.write(res.content)

        log(f"✅ 已下載到暫存：{full_path}")
        log(f"📦 檔案大小：{os.path.getsize(full_path)} bytes")
        log(f"📂 暫存目錄存在：{Path(tmpdir).exists()}")
        log(f"📄 暫存檔案存在：{Path(full_path).exists()}")

        # 本機存一份（本機跑才做）
        save_to_local_if_possible(full_path, filename)

        # 上傳 GDrive
        upload_to_gdrive(full_path, GDRIVE_FOLDER_ID)

        log(f"✅ 完成處理：{filename}")


def main():
    log("====================================================")
    log("schedule_report.py 開始執行")
    log(f"GITHUB_ACTIONS = {os.getenv('GITHUB_ACTIONS')}")
    log(f"IS_GITHUB = {IS_GITHUB}")
    log(f"PWD = {os.getcwd()}")
    log(f"LOCAL_OUTPUT_DIR = {LOCAL_OUTPUT_DIR}")
    log(f"LOCAL_OUTPUT_DIR exists = {LOCAL_OUTPUT_DIR.exists()}")
    log("====================================================")

    this_month, next_month, today_stamp, next_month_stamp = get_month_strings()
    log(f"this_month = {this_month}")
    log(f"next_month = {next_month}")
    log(f"today_stamp = {today_stamp}")
    log(f"next_month_stamp = {next_month_stamp}")

    regions = load_accounts()
    log(f"帳號數量 = {len(regions)}")

    for city, email, password in regions:
        log("")
        log(f"=== 開始處理 {city} ===")

        try:
            session = login(email, password)

            current_filename = f"排班統計表{today_stamp}-{city}.xlsx"
            next_filename = f"排班統計表{next_month_stamp}-{city}.xlsx"

            export_schedule(session, this_month, current_filename)
            export_schedule(session, next_month, next_filename)

            log(f"✅ {city} 全部完成")

        except Exception as e:
            log(f"❌ {city} 失敗：{e}")
            raise

    log("🎉 schedule_report.py 全部完成")


if __name__ == "__main__":
    main()
