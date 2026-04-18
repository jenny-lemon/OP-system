import os
import json
import calendar
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, date
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


def now_dt():
    return datetime.now()


def ensure_dirs():
    os.makedirs(DASHBOARD_DIR, exist_ok=True)
    os.makedirs(LATEST_DIR, exist_ok=True)
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    os.makedirs(EXEC_LOG_DIR, exist_ok=True)
    os.makedirs(DAILY_HISTORY_DIR, exist_ok=True)


def log(msg: str):
    print(f"[{now_dt().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def get_enabled_cities():
    enabled = [city for city in CITY_ORDER if city in ACCOUNTS]
    missing = [city for city in CITY_ORDER if city not in ACCOUNTS]

    if missing:
        log(f"⚠️ ACCOUNTS 缺少城市設定，已略過：{', '.join(missing)}")

    if not enabled:
        raise RuntimeError("ACCOUNTS 沒有任何可用城市設定")

    return enabled


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

    res = session.post(LOGIN_URL, data=payload, headers=HEADERS, allow_redirects=True)
    res.raise_for_status()

    if "login" in res.url.lower():
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
    name = str(name or "").strip()
    name = name.replace("螨", "蟎")

    mapping = {
        "VIP": "儲值金",
        "冷氣機清潔": "冷氣清潔",
        "冷氣機清潔服務": "冷氣清潔",
        "洗衣機": "洗衣機清潔",
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
        r"(\d{1,2}-\d{1,2})",
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
            if re.fullmatch(r"\d{1,2}-\d{1,2}", raw):
                today = datetime.today()
                dt = datetime.strptime(f"{today.year}-{raw}", "%Y-%m-%d")
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
            if not service:
                continue
            if service == "加總":
                continue
            if service.startswith("LC"):
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

                bm_paid = bm["已付款"].sum()
                bm_unpaid = bm["待付款"].sum()
                nm_paid = nm["已付款"].sum()
                nm_unpaid = nm["待付款"].sum()

                rows.append({
                    "城市": city,
                    "收入類型": income,
                    "類別": category,
                    "本月待付": bm_unpaid,
                    "本月已付": bm_paid,
                    "本月加總": bm_paid + bm_unpaid,
                    "次月待付": nm_unpaid,
                    "次月已付": nm_paid,
                    "次月加總": nm_paid + nm_unpaid,
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

        bm_total = total_df["本月加總"].sum()
        nm_total = total_df["次月加總"].sum()

        rows.append({
            "城市": city,
            "本月加總": bm_total,
            "次月加總": nm_total,
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


def build_daily_overview_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    cols = ["日期"]
    for city in CITY_ORDER:
        cols.extend([f"{city}業績", f"{city}佔比"])
    cols.append("全區合計")

    if raw_df.empty or "日期" not in raw_df.columns:
        return pd.DataFrame(columns=cols)

    work = raw_df.copy()
    work = work[work["月份"] == "本月"].copy()

    work["日期"] = pd.to_datetime(work["日期"], errors="coerce")
    work = work[work["日期"].notna()].copy()

    if work.empty:
        return pd.DataFrame(columns=cols)

    work["日期"] = work["日期"].dt.date
    work["業績"] = work["已付款"] + work["待付款"]

    today = datetime.today().date()
    first_day = date(today.year, today.month, 1)
    all_days = pd.date_range(first_day, today, freq="D").date

    grouped = work.groupby(["日期", "城市"], as_index=False)["業績"].sum()
    pivot = grouped.pivot(index="日期", columns="城市", values="業績").fillna(0)

    for city in CITY_ORDER:
        if city not in pivot.columns:
            pivot[city] = 0

    pivot = pivot[CITY_ORDER]
    pivot = pivot.reindex(all_days, fill_value=0)
    pivot["全區合計"] = pivot.sum(axis=1)

    out = pd.DataFrame(index=pivot.index)
    out["日期"] = [f"{d.month}/{d.day}" for d in pivot.index]

    for city in CITY_ORDER:
        out[f"{city}業績"] = pivot[city].astype(int)
        out[f"{city}佔比"] = pivot.apply(
            lambda r: 0 if r["全區合計"] == 0 else r[city] / r["全區合計"],
            axis=1
        )

    out["全區合計"] = pivot["全區合計"].astype(int)
    out = out.iloc[::-1].reset_index(drop=True)
    return out


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
            mail_df[col] = mail_df[col].apply(
                lambda x: f"{int(x):,}" if pd.notna(x) else ""
            )

    if "本月佔比" in mail_df.columns:
        mail_df["本月佔比"] = mail_df["本月佔比"].apply(
            lambda x: f"{x:.2%}" if pd.notna(x) else ""
        )

    if "次月佔比" in mail_df.columns:
        mail_df["次月佔比"] = mail_df["次月佔比"].apply(
            lambda x: f"{x:.2%}" if pd.notna(x) else ""
        )

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


def append_execution_log(df4: pd.DataFrame, trigger: str):
    ensure_dirs()

    now = now_dt()
    exec_file = os.path.join(EXEC_LOG_DIR, now.strftime("%Y%m") + ".csv")

    total_row = df4[df4["城市"] == "加總"]
    if total_row.empty:
        return

    total_row = total_row.iloc[0]

    row = {
        "id": now.strftime("%Y%m%d%H%M%S"),
        "執行時間": now.strftime("%Y-%m-%d %H:%M:%S"),
        "來源": trigger,
        "本月加總": int(total_row.get("本月加總", 0)),
        "次月加總": int(total_row.get("次月加總", 0)),
        "本月家電加總": int(total_row.get("本月家電加總", 0)),
        "次月家電加總": int(total_row.get("次月家電加總", 0)),
        "儲值金": int(total_row.get("儲值金", 0)),
    }

    new_df = pd.DataFrame([row])

    if os.path.exists(exec_file):
        old_df = pd.read_csv(exec_file)
        out_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        out_df = new_df

    out_df.to_csv(exec_file, index=False, encoding="utf-8-sig")


def load_execution_log_for_current_month() -> pd.DataFrame:
    now = now_dt()
    exec_file = os.path.join(EXEC_LOG_DIR, now.strftime("%Y%m") + ".csv")
    if not os.path.exists(exec_file):
        return pd.DataFrame(columns=[
            "id", "執行時間", "來源", "本月加總", "次月加總",
            "本月家電加總", "次月家電加總", "儲值金"
        ])
    return pd.read_csv(exec_file)


def append_daily_overview_history(daily_df: pd.DataFrame, trigger: str):
    ensure_dirs()

    now = now_dt()
    month_file = os.path.join(DAILY_HISTORY_DIR, now.strftime("%Y%m") + ".csv")

    total_today = 0
    if not daily_df.empty and "全區合計" in daily_df.columns:
        first_row = daily_df.iloc[0]
        total_today = int(first_row.get("全區合計", 0))

    row = {
        "id": now.strftime("%Y%m%d%H%M%S"),
        "執行時間": now.strftime("%Y-%m-%d %H:%M:%S"),
        "來源": trigger,
        "日期": now.strftime("%Y-%m-%d"),
        "今日全區合計": total_today,
        "daily_rows": int(len(daily_df)),
        "daily_json": daily_df.to_json(orient="records", force_ascii=False),
    }

    new_df = pd.DataFrame([row])

    if os.path.exists(month_file):
        old_df = pd.read_csv(month_file, encoding="utf-8-sig")
        out_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        out_df = new_df

    out_df.to_csv(month_file, index=False, encoding="utf-8-sig")


def load_daily_history_for_current_month() -> pd.DataFrame:
    now = now_dt()
    month_file = os.path.join(DAILY_HISTORY_DIR, now.strftime("%Y%m") + ".csv")
    if not os.path.exists(month_file):
        return pd.DataFrame(columns=[
            "id", "執行時間", "來源", "日期", "今日全區合計", "daily_rows", "daily_json"
        ])
    return pd.read_csv(month_file, encoding="utf-8-sig")


def delete_daily_history_rows(ids):
    if not ids:
        return 0

    now = now_dt()
    month_file = os.path.join(DAILY_HISTORY_DIR, now.strftime("%Y%m") + ".csv")
    if not os.path.exists(month_file):
        return 0

    df = pd.read_csv(month_file, encoding="utf-8-sig")
    before = len(df)
    df = df[~df["id"].astype(str).isin([str(x) for x in ids])].copy()
    df.to_csv(month_file, index=False, encoding="utf-8-sig")
    return before - len(df)


def persist_dashboard_payload(df4: pd.DataFrame, daily_df: pd.DataFrame, email_html: str):
    ensure_dirs()

    now = now_dt()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    day_folder = os.path.join(SNAPSHOT_DIR, now.strftime("%Y%m"))
    os.makedirs(day_folder, exist_ok=True)

    latest_df4 = os.path.join(LATEST_DIR, "df4.csv")
    latest_daily = os.path.join(LATEST_DIR, "daily_df.csv")
    latest_html = os.path.join(LATEST_DIR, "email_preview.html")
    latest_meta = os.path.join(LATEST_DIR, "meta.json")

    df4.to_csv(latest_df4, index=False, encoding="utf-8-sig")
    daily_df.to_csv(latest_daily, index=False, encoding="utf-8-sig")

    with open(latest_html, "w", encoding="utf-8") as f:
        f.write(email_html)

    meta = {
        "updated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "df4_rows": int(len(df4)),
        "daily_rows": int(len(daily_df)),
    }
    with open(latest_meta, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    snapshot_prefix = os.path.join(day_folder, stamp)
    df4.to_csv(f"{snapshot_prefix}_df4.csv", index=False, encoding="utf-8-sig")
    daily_df.to_csv(f"{snapshot_prefix}_daily_df.csv", index=False, encoding="utf-8-sig")

    with open(f"{snapshot_prefix}_meta.json", "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    with open(f"{snapshot_prefix}_email_preview.html", "w", encoding="utf-8") as f:
        f.write(email_html)


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
        return {
            "raw_df": pd.DataFrame(),
            "df1": pd.DataFrame(),
            "df2": pd.DataFrame(),
            "df3": pd.DataFrame(),
            "df4": pd.DataFrame(columns=[
                "城市", "本月加總", "本月佔比", "次月加總", "次月佔比",
                "本月家電加總", "次月家電加總", "儲值金"
            ]),
            "daily_df": pd.DataFrame(),
            "email_html": "",
            "updated_at": now_dt().strftime("%Y-%m-%d %H:%M:%S"),
            "execution_log_df": load_execution_log_for_current_month(),
            "daily_history_df": load_daily_history_for_current_month(),
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
                        res = session.get(url, headers=HEADERS, allow_redirects=True)
                        res.raise_for_status()

                        rows = parse_html(res.text)
                        city_row_count += len(rows)

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
        empty_daily = pd.DataFrame()

        if persist_dashboard:
            persist_dashboard_payload(empty_df4, empty_daily, "")

        return {
            "raw_df": pd.DataFrame(),
            "df1": pd.DataFrame(),
            "df2": pd.DataFrame(),
            "df3": pd.DataFrame(),
            "df4": empty_df4,
            "daily_df": empty_daily,
            "email_html": "",
            "updated_at": now_dt().strftime("%Y-%m-%d %H:%M:%S"),
            "execution_log_df": load_execution_log_for_current_month(),
            "daily_history_df": load_daily_history_for_current_month(),
            "error": error_msg,
        }

    df1 = build_region1_df(raw_df)
    df2 = build_region2_df(raw_df)
    df3 = build_region3_df(df2)
    df4 = build_region4_df(df2)
    daily_df = build_daily_overview_df(raw_df)
    email_html = build_region4_email_html(df4)

    append_execution_log(df4, trigger=trigger)
    append_daily_overview_history(daily_df, trigger=trigger)

    if persist_dashboard:
        persist_dashboard_payload(df4, daily_df, email_html)

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
        "execution_log_df": load_execution_log_for_current_month(),
        "daily_history_df": load_daily_history_for_current_month(),
        "error": None if not city_errors else " / ".join(city_errors),
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
