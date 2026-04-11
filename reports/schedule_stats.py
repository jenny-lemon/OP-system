from datetime import datetime

from gdrive import upload_bytes_to_gdrive

EXPORT_BASE = "https://backend.lemonclean.com.tw/schedule/export_times"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

# 你給的 Google Drive 資料夾 ID
FOLDER_ID = "1cNWX1eL6SzkjJH8qQQyoz6cUoFHoVrRq"


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


def export_schedule(session, month: str, filename: str) -> str:
    url = f"{EXPORT_BASE}?month={month}"
    res = session.get(url, headers=HEADERS, allow_redirects=True)

    content_type = (res.headers.get("Content-Type") or "").lower()
    if res.status_code != 200:
        raise RuntimeError(f"{filename} 下載失敗，status={res.status_code}")

    if "excel" not in content_type and "octet-stream" not in content_type:
        raise RuntimeError(f"{filename} 回傳不是 Excel，Content-Type={content_type}")

    file_id = upload_bytes_to_gdrive(
        content=res.content,
        filename=filename,
        folder_id=FOLDER_ID,
    )
    return file_id


def run_schedule_stats(session, city: str) -> list[dict]:
    this_month, next_month, today_stamp, next_month_stamp = get_month_strings()

    current_filename = f"排班統計表{today_stamp}-{city}.xlsx"
    next_filename = f"排班統計表{next_month_stamp}-{city}.xlsx"

    current_file_id = export_schedule(session, this_month, current_filename)
    next_file_id = export_schedule(session, next_month, next_filename)

    return [
        {"city": city, "filename": current_filename, "file_id": current_file_id},
        {"city": city, "filename": next_filename, "file_id": next_file_id},
    ]
