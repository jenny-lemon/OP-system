import os
import calendar
from datetime import datetime
import pytz

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build

LOGIN_URL = "https://backend.lemonclean.com.tw/login"
PURCHASE_URL = "https://backend.lemonclean.com.tw/purchase"

HEADERS = {"User-Agent": "Mozilla/5.0"}

CITY_ORDER = ["台北", "台中", "桃園", "新竹", "高雄"]

# 👉 台北時區
tz = pytz.timezone("Asia/Taipei")

# 👉 你的 Google Sheet ID
SHEET_ID = "👉填你的sheet id"
SHEET_NAME = "業績"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


# =========================
# 帳密（多區）
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
# Google Sheets
# =========================
def get_sheet_service():
    creds = service_account.Credentials.from_service_account_info(
        dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]),
        scopes=SCOPES,
    )
    return build("sheets", "v4", credentials=creds)


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
# 建 URL
# =========================
def build_url(start, end, status):
    params = {
        "clean_date_s": start,
        "clean_date_e": end,
        "purchase_status": str(status),
    }
    return requests.Request("GET", PURCHASE_URL, params=params).prepare().url


# =========================
# 解析（簡化版）
# =========================
def parse_html(html):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")

    total = 0

    for table in tables:
        rows = table.find_all("tr")
        for r in rows:
            cols = [c.get_text(strip=True) for c in r.find_all("td")]
            if len(cols) >= 3:
                try:
                    val = int(cols[-1].replace(",", ""))
                    total += val
                except:
                    pass

    return total


# =========================
# 寫入 Sheet
# =========================
def write_to_sheet(data_rows):
    service = get_sheet_service()

    body = {"values": data_rows}

    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=f"{SHEET_NAME}!A:H",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()

    print("📊 已寫入 Google Sheet")


# =========================
# 主程式
# =========================
def main():
    print("🔥 業績報表（Sheet版）")

    now = datetime.now(tz)
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")

    (m_start, m_end), (n_start, n_end) = get_ranges()

    accounts = load_accounts()
    results = []

    for city in CITY_ORDER:
        print(f"\n=== {city} ===")

        session = requests.Session()
        acc = accounts[city]

        try:
            login(session, acc["email"], acc["password"])

            # 本月
            url1 = build_url(m_start, m_end, 1)
            res1 = session.get(url1, headers=HEADERS)
            bm = parse_html(res1.text)

            # 次月
            url2 = build_url(n_start, n_end, 1)
            res2 = session.get(url2, headers=HEADERS)
            nm = parse_html(res2.text)

            results.append([date_str, time_str, city, bm, nm])

        except Exception as e:
            print(f"❌ {city} 失敗：{e}")

    df = pd.DataFrame(results, columns=["日期", "時間", "城市", "本月", "次月"])

    total = df["本月"].sum()
    df["佔比"] = df["本月"] / total if total else 0

    rows = df.values.tolist()

    write_to_sheet(rows)

    print("✅ 完成")


if __name__ == "__main__":
    main()
