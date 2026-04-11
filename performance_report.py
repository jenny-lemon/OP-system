import calendar
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build

# =========================
# 基本設定
# =========================
LOGIN_URL = "https://backend.lemonclean.com.tw/login"
PURCHASE_URL = "https://backend.lemonclean.com.tw/purchase"

HEADERS = {"User-Agent": "Mozilla/5.0"}

CITY_ORDER = ["台北", "台中", "桃園", "新竹", "高雄"]

# 👉 台北時區（不用 pytz）
tz = timezone(timedelta(hours=8))

# 👉 Google Drive 資料夾（放每月報表）
DRIVE_FOLDER_ID = "👉填你的資料夾ID"

# 權限
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
SHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


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
# Google API
# =========================
def get_creds(scopes):
    return service_account.Credentials.from_service_account_info(
        dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
        scopes=scopes,
    )


def get_drive():
    return build("drive", "v3", credentials=get_creds(DRIVE_SCOPES))


def get_sheets():
    return build("sheets", "v4", credentials=get_creds(SHEET_SCOPES))


# =========================
# 每月 Sheet
# =========================
def get_month_title():
    now = datetime.now(tz)
    return f"業績報表_{now.strftime('%Y-%m')}"


def find_sheet(drive, title):
    q = (
        f"name='{title}' and "
        f"mimeType='application/vnd.google-apps.spreadsheet' and "
        f"'{DRIVE_FOLDER_ID}' in parents and trashed=false"
    )

    res = drive.files().list(q=q, fields="files(id,name)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def create_sheet(drive, title):
    file = {
        "name": title,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [DRIVE_FOLDER_ID],
    }

    res = drive.files().create(body=file, fields="id,name").execute()
    print(f"🆕 建立：{title}")
    return res["id"]


def get_or_create_sheet():
    drive = get_drive()
    title = get_month_title()

    sid = find_sheet(drive, title)
    if sid:
        print(f"📄 使用：{title}")
        return sid

    return create_sheet(drive, title)


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
# API
# =========================
def build_url(start, end):
    params = {
        "clean_date_s": start,
        "clean_date_e": end,
        "purchase_status": "1",
    }
    return requests.Request("GET", PURCHASE_URL, params=params).prepare().url


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
# Sheet寫入
# =========================
def ensure_tab(service, sid, name):
    meta = service.spreadsheets().get(spreadsheetId=sid).execute()
    tabs = [s["properties"]["title"] for s in meta["sheets"]]

    if name in tabs:
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={"requests": [{"addSheet": {"properties": {"title": name}}}]},
    ).execute()


def write_sheet(rows):
    sid = get_or_create_sheet()
    service = get_sheets()

    now = datetime.now(tz)
    tab = now.strftime("%m-%d_%H%M")

    ensure_tab(service, sid, tab)

    service.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"{tab}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": rows},
    ).execute()

    print(f"📊 已寫入：{tab}")


# =========================
# 主程式
# =========================
def main():
    print("🔥 業績報表（Sheet版）")

    now = datetime.now(tz)
    date = now.strftime("%Y-%m-%d")
    time = now.strftime("%H:%M")

    (m_start, m_end), (n_start, n_end) = get_ranges()

    accounts = load_accounts()
    result = []

    for city in CITY_ORDER:
        print(f"\n=== {city} ===")

        session = requests.Session()
        acc = accounts[city]

        try:
            login(session, acc["email"], acc["password"])

            bm = parse_total(session.get(build_url(m_start, m_end)).text)
            nm = parse_total(session.get(build_url(n_start, n_end)).text)

            result.append([date, time, city, bm, nm])

        except Exception as e:
            print(f"❌ {city}：{e}")

    df = pd.DataFrame(result, columns=["日期", "時間", "城市", "本月", "次月"])

    total = df["本月"].sum()
    df["佔比"] = df["本月"] / total if total else 0

    rows = [df.columns.tolist()] + df.values.tolist()

    write_sheet(rows)

    print("✅ 完成")


if __name__ == "__main__":
    main()
