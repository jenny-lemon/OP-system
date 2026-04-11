import os
import calendar
import tempfile
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

# =========================
# 基本設定
# =========================
LOGIN_URL = "https://backend.lemonclean.com.tw/login"
PURCHASE_URL = "https://backend.lemonclean.com.tw/purchase"

HEADERS = {"User-Agent": "Mozilla/5.0"}

CITY_ORDER = ["台北", "台中", "桃園", "新竹", "高雄"]

# 👉 台北時區（不用 pytz）
tz = timezone(timedelta(hours=8))

# 👉 Google Drive 資料夾
DRIVE_FOLDER_ID = "1_b2EjuCAZ6qdlzUjY_PiecbpV7MKBC2t"

SCOPES = ["https://www.googleapis.com/auth/drive"]

# =========================
# 帳密
# =========================
def load_accounts():
    mapping = {
        "台北": "taipei",
        "台中": "taichung",
        "桃園": "taoyuan",
        "新竹": "hsinchu",
        "高雄": "kaohsiung",
    }

    accounts = {}
    for city, key in mapping.items():
        accounts[city] = {
            "email": st.secrets["accounts"][key]["email"],
            "password": st.secrets["accounts"][key]["password"],
        }
    return accounts


# =========================
# Google Drive
# =========================
def get_drive():
    creds = service_account.Credentials.from_service_account_info(
        dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
        scopes=SCOPES,
    )
    return build("drive", "v3", credentials=creds)


def find_file(drive, filename):
    q = f"name='{filename}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"

    res = drive.files().list(
        q=q,
        fields="files(id,name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = res.get("files", [])
    return files[0]["id"] if files else None


def download_file(drive, file_id, path):
    request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = open(path, "wb")

    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    fh.close()


def upload_file(drive, path, file_id=None):
    media = MediaFileUpload(path, resumable=True)

    if file_id:
        drive.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True,
        ).execute()
        print("☁️ 已更新 Excel")
    else:
        drive.files().create(
            body={"name": os.path.basename(path), "parents": [DRIVE_FOLDER_ID]},
            media_body=media,
            supportsAllDrives=True,
        ).execute()
        print("☁️ 已上傳 Excel")


# =========================
# 登入
# =========================
def login(session, email, password):
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

    print(f"✅ 登入成功：{email}")


# =========================
# 日期區間
# =========================
def get_ranges():
    now = datetime.now(tz)
    y, m = now.year, now.month

    this_start = f"{y}-{m:02d}-01"
    this_end = f"{y}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"

    if m == 12:
        ny, nm = y + 1, 1
    else:
        ny, nm = y, m + 1

    next_start = f"{ny}-{nm:02d}-01"
    next_end = f"{ny}-{nm:02d}-{calendar.monthrange(ny, nm)[1]:02d}"

    return (this_start, this_end), (next_start, next_end)


# =========================
# 抓數字（簡化版）
# =========================
def parse_total(html):
    soup = BeautifulSoup(html, "html.parser")
    total = 0

    for td in soup.find_all("td"):
        try:
            total += int(td.text.replace(",", ""))
        except:
            pass

    return total


# =========================
# 寫 Excel
# =========================
def write_excel(path, df, sheet_name):
    if os.path.exists(path):
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])

    ws = wb.create_sheet(sheet_name)

    ws.append(df.columns.tolist())

    for _, row in df.iterrows():
        ws.append(row.tolist())

    ws.freeze_panes = "A2"

    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.font = Font(bold=True)

    wb.save(path)


# =========================
# 主程式
# =========================
def main():
    print("🔥 業績報表（Excel版）")

    now = datetime.now(tz)

    sheet_name = now.strftime("%m-%d_%H%M")
    filename = f"業績報表_{now.strftime('%Y-%m')}.xlsx"

    (m_start, m_end), (n_start, n_end) = get_ranges()

    accounts = load_accounts()
    result = []

    for city in CITY_ORDER:
        print(f"\n=== {city} ===")

        session = requests.Session()
        acc = accounts[city]

        try:
            login(session, acc["email"], acc["password"])

            bm = parse_total(session.get(PURCHASE_URL).text)
            nm = parse_total(session.get(PURCHASE_URL).text)

            result.append([city, bm, nm])

        except Exception as e:
            print(f"❌ {city}：{e}")

    df = pd.DataFrame(result, columns=["城市", "本月", "次月"])

    drive = get_drive()

    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, filename)

        file_id = find_file(drive, filename)

        if file_id:
            download_file(drive, file_id, path)

        write_excel(path, df, sheet_name)

        upload_file(drive, path, file_id)

    print("✅ 完成")


if __name__ == "__main__":
    main()
