import os
import json
import calendar
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta, timezone

TZ_TAIPEI = timezone(timedelta(hours=8))

def now_dt():
    return datetime.now(TZ_TAIPEI)
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

from accounts import ACCOUNTS
from paths import PATH_REPORT


LOGIN_URL = "https://backend.lemonclean.com.tw/login"
PURCHASE_URL = "https://backend.lemonclean.com.tw/purchase"
HEADERS = {"User-Agent": "Mozilla/5.0"}

CITY_ORDER = ["台北", "台中", "桃園", "新竹", "高雄"]
INCOME_ORDER = ["現金收入", "儲值金"]
CATEGORY_ORDER = ["清潔", "儲值金", "冷氣", "洗衣機", "水洗", "收納"]

REGION3_CATEGORY_ORDER = [
    "清潔",
    "冷氣",
    "洗衣機",
    "水洗",
    "收納",
    "儲值金",
    "清潔現金+儲值金",
    "家電現金+儲值金",
    "水洗/收納現金+儲值金",
    "清潔+水洗+收納現金+儲值金",
]

DASHBOARD_DIR = os.path.join(PATH_REPORT, "_dashboard_sales")
LATEST_DIR = os.path.join(DASHBOARD_DIR, "latest")
SNAPSHOT_DIR = os.path.join(DASHBOARD_DIR, "snapshots")
EXEC_LOG_DIR = os.path.join(DASHBOARD_DIR, "execution_logs")
DAILY_HISTORY_DIR = os.path.join(DASHBOARD_DIR, "daily_overview_history")
OUTPUT_LOG_FILE = os.path.join(DASHBOARD_DIR, "output_file_log.csv")


def now_dt():
    return datetime.now()


def log(msg: str):
    print(f"[{now_dt().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def ensure_dirs():
    os.makedirs(DASHBOARD_DIR, exist_ok=True)
    os.makedirs(LATEST_DIR, exist_ok=True)
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    os.makedirs(EXEC_LOG_DIR, exist_ok=True)
    os.makedirs(DAILY_HISTORY_DIR, exist_ok=True)


def login(session, email, password):
    res = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, "html.parser")
    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise RuntimeError("找不到 _token，無法登入")

    payload = {
        "_token": token_input.get("value"),
        "email": email,
        "password": password,
    }

    login_res = session.post(LOGIN_URL, data=payload, headers=HEADERS, allow_redirects=True)
    login_res.raise_for_status()

    if "login" in login_res.url.lower():
        raise RuntimeError(f"{email} 登入失敗")

    log(f"✅ 登入成功：{email}")


def get_ranges():
    today = datetime.today()
    y, m = today.year, today.month

    this_start = f"{y}-{m:02d}-01"
    this_end = f"{y}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"

    if m == 12:
        ny, nm = y + 1, 1
    else:
        ny, nm = y, m + 1

    next_start = f"{ny}-{nm:02d}-01"
    next_end = f"{ny}-{nm:02d}-{calendar.monthrange(ny, nm)[1]:02d}"

    return (this_start, this_end), (next_start, next_end)


def build_url(start, end, status, keyword=""):
    params = {
        "keyword": keyword,
        "name": "",
        "phone": "",
        "orderNo": "",
        "date_s": "",
        "date_e": "",
        "clean_date_s": start,
        "clean_date_e": end,
        "paid_at_s": "",
        "paid_at_e": "",
        "refundDateS": "",
        "refundDateE": "",
        "buy": "",
        "area_id": "",
        "isCharge": "",
        "isRefund": "",
        "p_board": "on",
        "payway": "",
        "purchase_status": str(status),
        "progress_status": "",
        "invoiceStatus": "",
        "otherFee": "",
        "orderBy": "",
    }
    return requests.Request("GET", PURCHASE_URL, params=params).prepare().url


def get_keywords(city):
    if city == "新竹":
        return ["新竹"]
    if city == "高雄":
        return ["高雄", "台南"]
    return [""]


def safe_int(v):
    try:
        s = str(v).replace(",", "").strip()
        if s in ("", "-", "None", "nan"):
            return 0
        return int(float(s))
    except Exception:
        return 0


def normalize_service(name):
    name = str(name or "").strip().replace("螨", "蟎")

    mapping = {
        "VIP": "儲值金",
        "冷氣機清潔": "冷氣清潔",
        "冷氣機清潔服務": "冷氣清潔",
        "洗衣機": "洗衣機清潔",
        "洗衣機清潔": "洗衣機清潔",
        "沙發床墊水洗除蟎": "水洗",
        "沙發床墊水洗除螨": "水洗",
        "沙發清洗": "水洗",
        "床墊清洗": "水洗",
        "整理收納": "收納",
    }
    return mapping.get(name, name)


def detect_income_type(first_header):
    first_header = str(first_header or "").strip()
    if first_header in ("VIP", "儲值金"):
        return "儲值金"
    return "現金收入"


def normalize_date_text(text: str) -> Optional[str]:
    txt = str(text or "").strip()
    if not txt:
        return None

    txt = txt.replace("年", "-").replace("月", "-").replace("日", "")
    txt = txt.replace("/", "-").replace(".", "-")
    txt = " ".join(txt.split())

    import re

    patterns = [
        r"(20\d{2}-\d{1,2}-\d{1,2})",
        r"(20\d{6})",
        r"(\d{4}/\d{1,2}/\d{1,2})",
        r"(\d{4}\.\d{1,2}\.\d{1,2})",
        r"(\d{1,2}-\d{1,2})",
        r"(\d{1,2}/\d{1,2})",
    ]

    for p in patterns:
        m = re.search(p, txt)
        if not m:
            continue

        raw = m.group(1)
        try:
            if re.fullmatch(r"20\d{2}-\d{1,2}-\d{1,2}", raw):
                dt = datetime.strptime(raw, "%Y-%m-%d")
                return dt.strftime("%Y-%m-%d")
            if re.fullmatch(r"20\d{6}", raw):
                dt = datetime.strptime(raw, "%Y%m%d")
                return dt.strftime("%Y-%m-%d")
            if re.fullmatch(r"\d{4}/\d{1,2}/\d{1,2}", raw):
                dt = datetime.strptime(raw, "%Y/%m/%d")
                return dt.strftime("%Y-%m-%d")
            if re.fullmatch(r"\d{4}\.\d{1,2}\.\d{1,2}", raw):
                dt = datetime.strptime(raw, "%Y.%m.%d")
                return dt.strftime("%Y-%m-%d")
            if re.fullmatch(r"\d{1,2}-\d{1,2}", raw):
                today = datetime.today()
                dt = datetime.strptime(f"{today.year}-{raw}", "%Y-%m-%d")
                return dt.strftime("%Y-%m-%d")
            if re.fullmatch(r"\d{1,2}/\d{1,2}", raw):
                today = datetime.today()
                dt = datetime.strptime(f"{today.year}/{raw}", "%Y/%m/%d")
                return dt.strftime("%Y-%m-%d")
        except Exception:
            pass

    return None


def parse_html(html):
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    results = []

    date_candidates = ["服務日期", "清潔日期", "日期", "預約日期", "服務日", "clean_date"]

    for table in tables:
        trs = table.find_all("tr")
        rows = []

        for tr in trs:
            cells = tr.find_all(["th", "td"])
            row = [c.get_text(" ", strip=True) for c in cells]
            if any(str(x).strip() for x in row):
                rows.append(row)

        if not rows:
            continue

        header = [str(x).strip() for x in rows[0]]

        if "已付款金額" not in header and "待付款金額" not in header:
            continue

        paid_idx = header.index("已付款金額") if "已付款金額" in header else None
        unpaid_idx = header.index("待付款金額") if "待付款金額" in header else None

        date_idx = None
        for name in date_candidates:
            if name in header:
                date_idx = header.index(name)
                break

        income_type = detect_income_type(header[0] if header else "")
        source = "儲值金表" if income_type == "儲值金" else "主表"

        for row in rows[1:]:
            if not row:
                continue

            service = normalize_service(row[0] if len(row) > 0 else "")
            if not service or service == "加總" or service.startswith("LC"):
                continue

            paid = safe_int(row[paid_idx]) if paid_idx is not None and len(row) > paid_idx else 0
            unpaid = safe_int(row[unpaid_idx]) if unpaid_idx is not None and len(row) > unpaid_idx else 0

            service_date = None
            if date_idx is not None and len(row) > date_idx:
                service_date = normalize_date_text(row[date_idx])

            results.append({
                "收入類型": income_type,
                "資料來源": source,
                "服務": service,
                "子項目": "",
                "日期": service_date,
                "已付款": paid,
                "待付款": unpaid,
            })

    log(f"✅ parse_html rows = {len(results)}")
    return results


def to_category(service, income) -> Optional[str]:
    if service == "儲值金" and income == "現金收入":
        return "儲值金"
    if service in ["居家清潔", "辦公室清潔", "裝修細清", "搬入清潔", "搬出清潔", "大掃除"]:
        return "清潔"
    if service == "冷氣清潔":
        return "冷氣"
    if service == "洗衣機清潔":
        return "洗衣機"
    if service == "水洗":
        return "水洗"
    if service == "收納":
        return "收納"
    return None


def build_region1_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    work = raw_df.copy()

    work["本月已付款"] = 0
    work["本月待付款"] = 0
    work["下月已付款"] = 0
    work["下月待付款"] = 0

    this_mask = work["月份"] == "本月"
    next_mask = work["月份"] == "下月"

    work.loc[this_mask, "本月已付款"] = work.loc[this_mask, "已付款"]
    work.loc[this_mask, "本月待付款"] = work.loc[this_mask, "待付款"]
    work.loc[next_mask, "下月已付款"] = work.loc[next_mask, "已付款"]
    work.loc[next_mask, "下月待付款"] = work.loc[next_mask, "待付款"]

    region1 = (
        work.groupby(["城市", "收入類型", "資料來源", "服務", "子項目"], as_index=False)[
            ["本月已付款", "本月待付款", "下月已付款", "下月待付款"]
        ]
        .sum()
    )

    region1["城市"] = pd.Categorical(region1["城市"], categories=CITY_ORDER, ordered=True)
    region1["收入類型"] = pd.Categorical(region1["收入類型"], categories=INCOME_ORDER, ordered=True)
    region1 = region1.sort_values(["城市", "收入類型", "服務"]).reset_index(drop=True)

    region1["城市"] = region1["城市"].astype(str)
    region1["收入類型"] = region1["收入類型"].astype(str)
    return region1


def build_region2_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    work = raw_df.copy()
    work["類別"] = work.apply(lambda r: to_category(r["服務"], r["收入類型"]), axis=1)
    work = work[work["類別"].notna()].copy()

    rows = []
    for city in CITY_ORDER:
        for income in INCOME_ORDER:
            for category in CATEGORY_ORDER:
                sub = work[
                    (work["城市"] == city) &
                    (work["收入類型"] == income) &
                    (work["類別"] == category)
                ]

                bm = sub[sub["月份"] == "本月"]
                nm = sub[sub["月份"] == "下月"]

                rows.append({
                    "城市": city,
                    "收入類型": income,
                    "類別": category,
                    "本月待付": bm["待付款"].sum(),
                    "本月已付": bm["已付款"].sum(),
                    "本月加總": bm["已付款"].sum() + bm["待付款"].sum(),
                    "次月待付": nm["待付款"].sum(),
                    "次月已付": nm["已付款"].sum(),
                    "次月加總": nm["已付款"].sum() + nm["待付款"].sum(),
                })

    return pd.DataFrame(rows)


def build_region3_df(region2_df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for city in CITY_ORDER:
        city_df = region2_df[region2_df["城市"] == city].copy()

        level1 = city_df[["城市", "類別", "收入類型", "本月加總", "次月加總"]].copy()
        level1["加總類型"] = "加總1"
        rows.extend(level1.to_dict("records"))

        mapping_level2 = {
            "清潔現金+儲值金": ["清潔"],
            "家電現金+儲值金": ["冷氣", "洗衣機"],
            "水洗/收納現金+儲值金": ["水洗", "收納"],
        }

        for new_cat, old_cats in mapping_level2.items():
            tmp = city_df[city_df["類別"].isin(old_cats)]
            rows.append({
                "城市": city,
                "類別": new_cat,
                "收入類型": "現金+儲值金",
                "本月加總": tmp["本月加總"].sum(),
                "次月加總": tmp["次月加總"].sum(),
                "加總類型": "加總2",
            })

        mapping_level3 = {
            "清潔+水洗+收納現金+儲值金": ["清潔", "水洗", "收納"],
            "家電現金+儲值金": ["冷氣", "洗衣機"],
        }

        for new_cat, old_cats in mapping_level3.items():
            tmp = city_df[city_df["類別"].isin(old_cats)]
            rows.append({
                "城市": city,
                "類別": new_cat,
                "收入類型": "現金+儲值金",
                "本月加總": tmp["本月加總"].sum(),
                "次月加總": tmp["次月加總"].sum(),
                "加總類型": "加總3",
            })

    region3 = pd.DataFrame(rows)
    type_order = ["加總1", "加總2", "加總3"]

    region3["城市"] = pd.Categorical(region3["城市"], categories=CITY_ORDER, ordered=True)
    region3["加總類型"] = pd.Categorical(region3["加總類型"], categories=type_order, ordered=True)
    region3["類別"] = pd.Categorical(region3["類別"], categories=REGION3_CATEGORY_ORDER, ordered=True)

    region3 = region3.sort_values(["城市", "加總類型", "類別", "收入類型"]).reset_index(drop=True)
    region3["城市"] = region3["城市"].astype(str)
    region3["類別"] = region3["類別"].astype(str)
    region3["收入類型"] = region3["收入類型"].astype(str)
    region3["加總類型"] = region3["加總類型"].astype(str)
    return region3


def build_region4_df(region2_df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for city in CITY_ORDER:
        city_df = region2_df[region2_df["城市"] == city].copy()

        appliance_df = city_df[city_df["類別"].isin(["冷氣", "洗衣機"])]
        bm_appliance = appliance_df["本月加總"].sum()
        nm_appliance = appliance_df["次月加總"].sum()

        cash_stored_df = city_df[
            (city_df["收入類型"] == "現金收入") &
            (city_df["類別"] == "儲值金")
        ]
        bm_cash_stored = cash_stored_df["本月加總"].sum()
        nm_cash_stored = cash_stored_df["次月加總"].sum()

        total_df = city_df[
            ~city_df["類別"].isin(["冷氣", "洗衣機"]) &
            ~(
                (city_df["收入類型"] == "現金收入") &
                (city_df["類別"] == "儲值金")
            )
        ]

        rows.append({
            "城市": city,
            "本月加總": total_df["本月加總"].sum(),
            "次月加總": total_df["次月加總"].sum(),
            "本月家電加總": bm_appliance,
            "次月家電加總": nm_appliance,
            "儲值金": bm_cash_stored + nm_cash_stored,
        })

    region4 = pd.DataFrame(rows)
    bm_sum = region4["本月加總"].sum()
    nm_sum = region4["次月加總"].sum()

    region4["本月佔比"] = 0 if bm_sum == 0 else region4["本月加總"] / bm_sum
    region4["次月佔比"] = 0 if nm_sum == 0 else region4["次月加總"] / nm_sum

    total_row = pd.DataFrame([{
        "城市": "加總",
        "本月加總": bm_sum,
        "本月佔比": 1,
        "次月加總": nm_sum,
        "次月佔比": 1,
        "本月家電加總": region4["本月家電加總"].sum(),
        "次月家電加總": region4["次月家電加總"].sum(),
        "儲值金": region4["儲值金"].sum(),
    }])

    region4 = pd.concat([region4, total_row], ignore_index=True)

    return region4[[
        "城市",
        "本月加總",
        "本月佔比",
        "次月加總",
        "次月佔比",
        "本月家電加總",
        "次月家電加總",
        "儲值金",
    ]]


def build_daily_overview_df(df4: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "id",
        "來源",
        "日期",
        "台北業績", "台北佔比",
        "台中業績", "台中佔比",
        "桃園業績", "桃園佔比",
        "新竹業績", "新竹佔比",
        "高雄業績", "高雄佔比",
        "全區合計",
    ]

    if df4 is None or df4.empty:
        log("⚠️ build_daily_overview_df：df4 為空")
        return pd.DataFrame(columns=cols)

    latest_daily = os.path.join(LATEST_DIR, "daily_df.csv")
    now_obj = now_dt()   # 這裡 now_dt() 要是台北時區版本
    row_id = now_obj.strftime("%Y%m%d%H%M%S")
    date_text = now_obj.strftime("%Y/%m/%d %H:%M")

    def get_val(city, col):
        try:
            row = df4[df4["城市"] == city]
            if row.empty:
                return 0
            return row.iloc[0][col]
        except Exception:
            return 0

    if os.path.exists(latest_daily):
        try:
            old_df = pd.read_csv(latest_daily, encoding="utf-8-sig")
        except Exception:
            old_df = pd.DataFrame(columns=cols)
    else:
        old_df = pd.DataFrame(columns=cols)

    for c in cols:
        if c not in old_df.columns:
            old_df[c] = ""

    new_row = {
        "id": row_id,
        "來源": "dashboard",
        "日期": date_text,
        "台北業績": get_val("台北", "本月加總"),
        "台北佔比": get_val("台北", "本月佔比"),
        "台中業績": get_val("台中", "本月加總"),
        "台中佔比": get_val("台中", "本月佔比"),
        "桃園業績": get_val("桃園", "本月加總"),
        "桃園佔比": get_val("桃園", "本月佔比"),
        "新竹業績": get_val("新竹", "本月加總"),
        "新竹佔比": get_val("新竹", "本月佔比"),
        "高雄業績": get_val("高雄", "本月加總"),
        "高雄佔比": get_val("高雄", "本月佔比"),
        "全區合計": get_val("加總", "本月加總"),
    }

    out = pd.concat([old_df[cols], pd.DataFrame([new_row])], ignore_index=True)

    # 只保留台北時間當月資料
    current_prefix = now_obj.strftime("%Y/%m")
    out = out[out["日期"].astype(str).str.startswith(current_prefix)].copy()

    # 轉成 datetime 後排序，最新在最上面
    out["_sort_dt"] = pd.to_datetime(out["日期"], format="%Y/%m/%d %H:%M", errors="coerce")
    out = out.sort_values(["_sort_dt", "id"], ascending=[False, False]).drop(columns=["_sort_dt"])

    out = out.reset_index(drop=True)

    log(f"✅ build_daily_overview_df 完成，筆數 = {len(out)}")
    return out[cols]


def format_region4_for_display(df4: pd.DataFrame) -> pd.DataFrame:
    out = df4.copy()
    for col in ["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"]:
        if col in out.columns:
            out[col] = out[col].apply(lambda x: int(x) if pd.notna(x) else 0)
    return out


def build_region4_email_html(df4):
    mail_df = df4.copy()

    for col in ["本月加總", "次月加總", "本月家電加總", "次月家電加總", "儲值金"]:
        if col in mail_df.columns:
            mail_df[col] = mail_df[col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")

    if "本月佔比" in mail_df.columns:
        mail_df["本月佔比"] = mail_df["本月佔比"].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "")

    if "次月佔比" in mail_df.columns:
        mail_df["次月佔比"] = mail_df["次月佔比"].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "")

    html_table = mail_df.to_html(index=False, border=0)

    return f"""
    <html>
      <head>
        <style>
          table {{
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 14px;
          }}
          th, td {{
            border: 1px solid #999;
            padding: 6px 10px;
          }}
          th {{
            background-color: #f2f2f2;
            text-align: center;
          }}
          td {{
            text-align: right;
          }}
          td:first-child {{
            text-align: left;
          }}
        </style>
      </head>
      <body>
        <p>您好，以下為業績報表：</p>
        {html_table}
      </body>
    </html>
    """


def send_region4_email(df4, recipient="jenny@lemonclean.com.tw"):
    sender = "jenny@lemonclean.com.tw"
    password = "bkhe akob wvse ibhm"

    today_str = datetime.today().strftime("%Y%m%d")
    subject = f"業績報表{today_str}"
    html = build_region4_email_html(df4)

    msg = MIMEText(html, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, [recipient], msg.as_string())

    log(f"✅ 已寄出：{recipient}")


def load_execution_log_for_current_month() -> pd.DataFrame:
    return pd.DataFrame()


def delete_execution_log_rows(ids):
    return 0


def append_daily_overview_history(daily_df: pd.DataFrame, trigger: str):
    return None


def load_daily_history_for_current_month() -> pd.DataFrame:
    return pd.DataFrame()


def delete_daily_history_rows(ids):
    return 0


def append_output_file_log(category: str, file_path: str, trigger: str):
    ensure_dirs()

    row = {
        "id": now_dt().strftime("%Y%m%d%H%M%S%f"),
        "時間": now_dt().strftime("%Y-%m-%d %H:%M:%S"),
        "分類": category,
        "檔名": os.path.basename(file_path),
        "完整路徑": file_path,
        "trigger": trigger,
    }

    new_df = pd.DataFrame([row])

    if os.path.exists(OUTPUT_LOG_FILE):
        old_df = pd.read_csv(OUTPUT_LOG_FILE, encoding="utf-8-sig")
        out_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        out_df = new_df

    out_df.to_csv(OUTPUT_LOG_FILE, index=False, encoding="utf-8-sig")


def load_output_file_log() -> pd.DataFrame:
    ensure_dirs()
    if not os.path.exists(OUTPUT_LOG_FILE):
        return pd.DataFrame(columns=["id", "時間", "分類", "檔名", "完整路徑", "trigger"])
    return pd.read_csv(OUTPUT_LOG_FILE, encoding="utf-8-sig")


def persist_dashboard_payload(
    df4: pd.DataFrame,
    daily_df: pd.DataFrame,
    email_html: str,
    error_msg: Optional[str] = None,
    trigger: str = "dashboard",
):
    ensure_dirs()

    now = now_dt()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    month_folder = os.path.join(SNAPSHOT_DIR, now.strftime("%Y%m"))
    os.makedirs(month_folder, exist_ok=True)

    latest_df4 = os.path.join(LATEST_DIR, "df4.csv")
    latest_daily = os.path.join(LATEST_DIR, "daily_df.csv")
    latest_html = os.path.join(LATEST_DIR, "email_preview.html")
    latest_meta = os.path.join(LATEST_DIR, "meta.json")

    df4.to_csv(latest_df4, index=False, encoding="utf-8-sig")
    append_output_file_log("業績報表", latest_df4, trigger)

    daily_df.to_csv(latest_daily, index=False, encoding="utf-8-sig")
    append_output_file_log("業績報表", latest_daily, trigger)

    with open(latest_html, "w", encoding="utf-8") as f:
        f.write(email_html or "")
    append_output_file_log("業績報表", latest_html, trigger)

    meta = {
        "updated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "df4_rows": int(len(df4)),
        "daily_rows": int(len(daily_df)),
        "error": error_msg,
    }
    with open(latest_meta, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    append_output_file_log("業績報表", latest_meta, trigger)

    snapshot_prefix = os.path.join(month_folder, stamp)

    snap_df4 = f"{snapshot_prefix}_df4.csv"
    snap_daily = f"{snapshot_prefix}_daily_df.csv"
    snap_meta = f"{snapshot_prefix}_meta.json"
    snap_html = f"{snapshot_prefix}_email_preview.html"

    df4.to_csv(snap_df4, index=False, encoding="utf-8-sig")
    append_output_file_log("業績報表", snap_df4, trigger)

    daily_df.to_csv(snap_daily, index=False, encoding="utf-8-sig")
    append_output_file_log("業績報表", snap_daily, trigger)

    with open(snap_meta, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    append_output_file_log("業績報表", snap_meta, trigger)

    with open(snap_html, "w", encoding="utf-8") as f:
        f.write(email_html or "")
    append_output_file_log("業績報表", snap_html, trigger)


def generate_sales_report(send_email=False, persist_dashboard=True, trigger="dashboard"):
    log("🔥 開始業績報表")

    ensure_dirs()
    (m_start, m_end), (n_start, n_end) = get_ranges()
    merged = {}
    city_errors = []

    enabled_cities = [city for city in CITY_ORDER if city in ACCOUNTS]
    missing_cities = [city for city in CITY_ORDER if city not in ACCOUNTS]

    if missing_cities:
        log(f"⚠️ ACCOUNTS 缺少城市設定，已略過：{', '.join(missing_cities)}")

    if not enabled_cities:
        error_msg = "ACCOUNTS 沒有任何可用城市設定"
        log(f"❌ {error_msg}")

        empty_df4 = pd.DataFrame(columns=[
            "城市", "本月加總", "本月佔比", "次月加總", "次月佔比",
            "本月家電加總", "次月家電加總", "儲值金"
        ])
        empty_daily = pd.DataFrame(columns=[
            "id", "來源", "日期",
            "台北業績", "台北佔比",
            "台中業績", "台中佔比",
            "桃園業績", "桃園佔比",
            "新竹業績", "新竹佔比",
            "高雄業績", "高雄佔比",
            "全區合計",
        ])

        if persist_dashboard:
            persist_dashboard_payload(empty_df4, empty_daily, "", error_msg, trigger=trigger)

        return {
            "raw_df": pd.DataFrame(),
            "df1": pd.DataFrame(),
            "df2": pd.DataFrame(),
            "df3": pd.DataFrame(),
            "df4": empty_df4,
            "daily_df": empty_daily,
            "email_html": "",
            "updated_at": now_dt().strftime("%Y-%m-%d %H:%M:%S"),
            "execution_log_df": pd.DataFrame(),
            "daily_history_df": pd.DataFrame(),
            "output_file_log_df": load_output_file_log(),
            "error": error_msg,
        }

    for city in enabled_cities:
        log(f"===== {city} =====")
        session = requests.Session()
        acc = ACCOUNTS[city]

        try:
            login(session, acc["email"], acc["password"])
            city_row_count = 0

            for label, (s, e) in {
                "本月": (m_start, m_end),
                "下月": (n_start, n_end),
            }.items():
                for status in [1, 0]:
                    for kw in get_keywords(city):
                        url = build_url(s, e, status, kw)
                        log(f"抓取：city={city} month={label} status={status} kw={kw} url={url}")

                        res = session.get(url, headers=HEADERS, allow_redirects=True)
                        res.raise_for_status()

                        rows = parse_html(res.text)
                        city_row_count += len(rows)

                        if not rows:
                            log(f"⚠️ {city} / {label} / status={status} / kw={kw} 沒抓到資料，HTML 長度={len(res.text)}")
                            try:
                                debug_dir = os.path.join(DASHBOARD_DIR, "_debug_html")
                                os.makedirs(debug_dir, exist_ok=True)
                                debug_name = f"{city}_{label}_status{status}_{(kw or 'ALL')}.html"
                                debug_path = os.path.join(debug_dir, debug_name)
                                with open(debug_path, "w", encoding="utf-8") as f:
                                    f.write(res.text)
                                log(f"📝 已輸出 debug html：{debug_path}")
                                append_output_file_log("業績報表", debug_path, trigger)
                            except Exception as dbg_e:
                                log(f"⚠️ debug html 寫出失敗：{dbg_e}")

                        for row in rows:
                            key = (
                                city,
                                label,
                                row["日期"],
                                row["收入類型"],
                                row["資料來源"],
                                row["服務"],
                                row["子項目"],
                            )

                            if key not in merged:
                                merged[key] = {
                                    "城市": city,
                                    "月份": label,
                                    "日期": row["日期"],
                                    "收入類型": row["收入類型"],
                                    "資料來源": row["資料來源"],
                                    "服務": row["服務"],
                                    "子項目": row["子項目"],
                                    "已付款": 0,
                                    "待付款": 0,
                                }

                            merged[key]["已付款"] += row["已付款"]
                            merged[key]["待付款"] += row["待付款"]

            if city_row_count == 0:
                msg = f"{city}：登入成功，但沒有抓到任何表格資料"
                city_errors.append(msg)
                log(f"⚠️ {msg}")

        except Exception as e:
            msg = f"{city} 失敗：{e}"
            city_errors.append(msg)
            log(f"❌ {msg}")

    raw_df = pd.DataFrame(merged.values())

    if raw_df.empty:
        error_msg = "沒有任何資料可輸出"
        if city_errors:
            error_msg += "；" + " / ".join(city_errors)

        log(f"⚠️ {error_msg}")

        empty_df4 = pd.DataFrame(columns=[
            "城市", "本月加總", "本月佔比", "次月加總", "次月佔比",
            "本月家電加總", "次月家電加總", "儲值金"
        ])
        empty_daily = pd.DataFrame(columns=[
            "id", "來源", "日期",
            "台北業績", "台北佔比",
            "台中業績", "台中佔比",
            "桃園業績", "桃園佔比",
            "新竹業績", "新竹佔比",
            "高雄業績", "高雄佔比",
            "全區合計",
        ])

        if persist_dashboard:
            persist_dashboard_payload(empty_df4, empty_daily, "", error_msg, trigger=trigger)

        return {
            "raw_df": pd.DataFrame(),
            "df1": pd.DataFrame(),
            "df2": pd.DataFrame(),
            "df3": pd.DataFrame(),
            "df4": empty_df4,
            "daily_df": empty_daily,
            "email_html": "",
            "updated_at": now_dt().strftime("%Y-%m-%d %H:%M:%S"),
            "execution_log_df": pd.DataFrame(),
            "daily_history_df": pd.DataFrame(),
            "output_file_log_df": load_output_file_log(),
            "error": error_msg,
        }

    df1 = build_region1_df(raw_df)
    df2 = build_region2_df(raw_df)
    df3 = build_region3_df(df2)
    df4 = build_region4_df(df2)

    daily_df = build_daily_overview_df(df4)
    if not daily_df.empty and "來源" in daily_df.columns:
        daily_df.loc[daily_df.index[-1], "來源"] = trigger

    log(f"raw_df columns = {list(raw_df.columns)}")
    log(f"raw_df 前5筆 = {raw_df.head().to_dict('records')}")
    log(f"df1 rows = {len(df1)}")
    log(f"df2 rows = {len(df2)}")
    log(f"df3 rows = {len(df3)}")
    log(f"df4 rows = {len(df4)}")
    log(f"daily_df rows = {len(daily_df)}")

    email_html = build_region4_email_html(df4)
    error_msg = None if not city_errors else " / ".join(city_errors)

    if persist_dashboard:
        persist_dashboard_payload(df4, daily_df, email_html, error_msg, trigger=trigger)

    if send_email:
        send_region4_email(df4)

    return {
        "raw_df": raw_df,
        "df1": df1,
        "df2": df2,
        "df3": df3,
        "df4": format_region4_for_display(df4),
        "daily_df": daily_df,
        "email_html": email_html,
        "updated_at": now_dt().strftime("%Y-%m-%d %H:%M:%S"),
        "execution_log_df": pd.DataFrame(),
        "daily_history_df": pd.DataFrame(),
        "output_file_log_df": load_output_file_log(),
        "error": error_msg,
    }


def main():
    trigger = "schedule"
    send_email = True

    if len(os.sys.argv) >= 2:
        trigger = os.sys.argv[1]

    if len(os.sys.argv) >= 3:
        send_email = os.sys.argv[2].lower() in ("1", "true", "yes", "y")

    generate_sales_report(
        send_email=send_email,
        persist_dashboard=True,
        trigger=trigger,
    )


if __name__ == "__main__":
    main()
