"""
performance_report.py

功能：
1. 產生 df4 / daily_df / email_html
2. dashboard 執行或排程執行時，累積「當月每日業績總覽」歷史紀錄
3. 記錄輸出檔案完整路徑
4. 提供 dashboard_main.py 需要的載入 / 刪除函式
"""

import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Any, List

import pandas as pd

TZ_TAIPEI = timezone(timedelta(hours=8))
BASE_DIR = Path("/Users/jenny/lemon") if Path("/Users/jenny/lemon").exists() else Path(__file__).resolve().parent

LATEST_DIR = BASE_DIR / "latest"
LATEST_DIR.mkdir(parents=True, exist_ok=True)

DAILY_OVERVIEW_HISTORY_CSV = LATEST_DIR / "daily_overview_history.csv"
OUTPUT_FILE_LOG_CSV = LATEST_DIR / "output_file_log.csv"
META_JSON = LATEST_DIR / "meta.json"
DF4_CSV = LATEST_DIR / "df4.csv"
DAILY_DF_CSV = LATEST_DIR / "daily_df.csv"
EMAIL_PREVIEW_HTML = LATEST_DIR / "email_preview.html"


# ── 基本工具 ──────────────────────────────────────────────────────────────────

def now_dt() -> datetime:
    return datetime.now(TZ_TAIPEI)

def now_str() -> str:
    return now_dt().strftime("%Y-%m-%d %H:%M:%S")

def today_ym() -> str:
    return now_dt().strftime("%Y-%m")

def _safe_read_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        return pd.read_csv(path)

def _safe_write_csv(df: pd.DataFrame, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")

def _ensure_id_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "id" not in df.columns:
        df.insert(0, "id", range(1, len(df) + 1))
    return df

def _reset_id_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy().reset_index(drop=True)
    if "id" in df.columns:
        df = df.drop(columns=["id"])
    df.insert(0, "id", range(1, len(df) + 1))
    return df


# ── 輸出檔案路徑紀錄 ──────────────────────────────────────────────────────────

def append_output_file_log(category: str, file_path: str, trigger: str = ""):
    p = Path(file_path)
    row = {
        "id": None,
        "分類": category,
        "檔名": p.name,
        "完整路徑": str(p),
        "大小": p.stat().st_size if p.exists() else 0,
        "建立時間": now_str(),
        "trigger": trigger,
    }

    df = load_output_file_log()
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df = _reset_id_column(df)
    _safe_write_csv(df, OUTPUT_FILE_LOG_CSV)

def load_output_file_log() -> pd.DataFrame:
    df = _safe_read_csv(OUTPUT_FILE_LOG_CSV)
    if df.empty:
        return pd.DataFrame(columns=["id", "分類", "檔名", "完整路徑", "大小", "建立時間", "trigger"])
    return _ensure_id_column(df)


# ── 每日業績總覽累積紀錄 ──────────────────────────────────────────────────────

def _build_daily_overview_history_rows(daily_df: pd.DataFrame, trigger: str) -> pd.DataFrame:
    if daily_df is None or daily_df.empty:
        return pd.DataFrame()

    rows = []
    executed_at = now_str()
    ym = today_ym()

    for _, row in daily_df.iterrows():
        item = {"id": None, "月份": ym, "執行時間": executed_at, "trigger": trigger}
        for col in daily_df.columns:
            item[col] = row[col]
        rows.append(item)

    return pd.DataFrame(rows)

def append_daily_overview_history(daily_df: pd.DataFrame, trigger: str = "dashboard") -> pd.DataFrame:
    new_rows = _build_daily_overview_history_rows(daily_df, trigger)
    old_df = load_daily_overview_history()

    if new_rows.empty:
        return old_df

    merged = pd.concat([old_df, new_rows], ignore_index=True)
    merged = _reset_id_column(merged)
    _safe_write_csv(merged, DAILY_OVERVIEW_HISTORY_CSV)
    return merged

def load_daily_overview_history() -> pd.DataFrame:
    df = _safe_read_csv(DAILY_OVERVIEW_HISTORY_CSV)
    if df.empty:
        return pd.DataFrame(columns=["id", "月份", "執行時間", "trigger"])
    return _ensure_id_column(df)

def load_daily_overview_history_for_current_month() -> pd.DataFrame:
    df = load_daily_overview_history()
    if df.empty:
        return df
    if "月份" not in df.columns:
        return df
    return df[df["月份"].astype(str) == today_ym()].reset_index(drop=True)

def delete_daily_overview_history_rows(ids: List[str]) -> int:
    if not ids:
        return 0
    df = load_daily_overview_history()
    if df.empty:
        return 0

    before = len(df)
    ids_set = {str(x) for x in ids}
    df = df[~df["id"].astype(str).isin(ids_set)].copy()
    deleted = before - len(df)

    if len(df) == 0:
        _safe_write_csv(pd.DataFrame(columns=["id", "月份", "執行時間", "trigger"]), DAILY_OVERVIEW_HISTORY_CSV)
    else:
        df = _reset_id_column(df)
        _safe_write_csv(df, DAILY_OVERVIEW_HISTORY_CSV)

    return deleted


# ── 舊函式保留，避免 dashboard_main 或其他舊程式炸掉 ────────────────────────

def load_execution_log_for_current_month() -> pd.DataFrame:
    # 舊版 dashboard_main 會叫這個；現在直接回傳空表，避免報錯
    return pd.DataFrame()

def delete_execution_log_rows(ids: List[str]) -> int:
    # 舊版 dashboard_main 會叫這個；現在直接不處理
    return 0


# ── 你的原本業績資料邏輯請貼在這裡 ───────────────────────────────────────────

def _run_existing_sales_pipeline(send_email: bool = False) -> Dict[str, Any]:
    """
    請把你原本 performance_report.py 裡面真正產出 df4 / daily_df / email_html
    的邏輯貼進這裡。

    你至少要回傳：
    {
        "df4": pd.DataFrame(...),
        "daily_df": pd.DataFrame(...),
        "email_html": "<html>...</html>",
        "error": None 或錯誤訊息字串
    }
    """

    # ===== 下面是示範格式，請換成你原本的邏輯 =====
    df4 = pd.DataFrame(columns=["城市", "本月加總", "本月佔比", "次月加總", "次月佔比", "本月家電加總", "次月家電加總", "儲值金"])
    daily_df = pd.DataFrame(columns=["日期", "台北業績", "台北佔比", "台中業績", "台中佔比", "桃園業績", "桃園佔比", "新竹業績", "新竹佔比", "高雄業績", "高雄佔比", "全區合計"])
    email_html = ""

    return {
        "df4": df4,
        "daily_df": daily_df,
        "email_html": email_html,
        "error": None,
    }


# ── 寄信 ──────────────────────────────────────────────────────────────────────

def send_region4_email(df4: pd.DataFrame):
    """
    這裡請保留你原本的寄信邏輯。
    如果你原本已有 send_region4_email，請直接把原本函式內容貼回來。
    """
    # 沒有原始碼的情況下，先保留空實作，避免 dashboard_main import 失敗
    return True


# ── 主流程 ────────────────────────────────────────────────────────────────────

def generate_sales_report(send_email: bool = False, persist_dashboard: bool = True, trigger: str = "dashboard") -> Dict[str, Any]:
    try:
        result = _run_existing_sales_pipeline(send_email=send_email)

        df4 = result.get("df4", pd.DataFrame())
        daily_df = result.get("daily_df", pd.DataFrame())
        email_html = result.get("email_html", "")
        error_msg = result.get("error")

        if persist_dashboard:
            if isinstance(df4, pd.DataFrame):
                _safe_write_csv(df4, DF4_CSV)
                append_output_file_log("業績報表", str(DF4_CSV), trigger=trigger)

            if isinstance(daily_df, pd.DataFrame):
                _safe_write_csv(daily_df, DAILY_DF_CSV)
                append_output_file_log("業績報表", str(DAILY_DF_CSV), trigger=trigger)

            if isinstance(email_html, str):
                EMAIL_PREVIEW_HTML.write_text(email_html, encoding="utf-8")
                append_output_file_log("業績報表", str(EMAIL_PREVIEW_HTML), trigger=trigger)

            meta = {
                "updated_at": now_str(),
                "trigger": trigger,
                "error": error_msg,
            }
            META_JSON.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
            append_output_file_log("業績報表", str(META_JSON), trigger=trigger)

        # dashboard 手動更新、或排程跑 performance_report.py dashboard false
        # 都要累積每日業績總覽
        daily_overview_history_df = append_daily_overview_history(daily_df, trigger=trigger)

        if send_email and not df4.empty:
            send_region4_email(df4)

        return {
            "df4": df4,
            "daily_df": daily_df,
            "email_html": email_html,
            "daily_overview_history_df": load_daily_overview_history_for_current_month(),
            "error": error_msg,
        }

    except Exception as e:
        meta = {
            "updated_at": now_str(),
            "trigger": trigger,
            "error": str(e),
        }
        META_JSON.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        return {
            "df4": pd.DataFrame(),
            "daily_df": pd.DataFrame(),
            "email_html": "",
            "daily_overview_history_df": load_daily_overview_history_for_current_month(),
            "error": str(e),
        }


# ── CLI ───────────────────────────────────────────────────────────────────────

def _parse_bool_arg(v: str) -> bool:
    return str(v).strip().lower() in {"1", "true", "yes", "y"}

if __name__ == "__main__":
    # 支援：
    # python performance_report.py
    # python performance_report.py dashboard false
    # python performance_report.py dashboard true
    args = sys.argv[1:]

    trigger = "schedule"
    send_email_flag = False

    if len(args) >= 1:
        trigger = args[0] or "schedule"
    if len(args) >= 2:
        send_email_flag = _parse_bool_arg(args[1])

    result = generate_sales_report(
        send_email=send_email_flag,
        persist_dashboard=True,
        trigger=trigger,
    )

    if result.get("error"):
        print(f"❌ {result['error']}")
        sys.exit(1)

    print("✅ performance_report 完成")
    sys.exit(0)
