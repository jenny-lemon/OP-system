from pathlib import Path
import os

# 是否在 GitHub Actions 執行
IS_GITHUB = os.getenv("GITHUB_ACTIONS") == "true"

if IS_GITHUB:
    # GitHub runner 用的暫存輸出根目錄
    BASE_OUTPUT = Path(os.getenv("OUTPUT_BASE", "/home/runner/work/lemon_outputs"))
    BASE_OUTPUT.mkdir(parents=True, exist_ok=True)

    PATH_JENNY = BASE_OUTPUT / "Jenny"
    PATH_SCHEDULE = BASE_OUTPUT / "排班統計表"
    PATH_CLEANER_SCHEDULE = BASE_OUTPUT / "專員班表"
    PATH_CLEANER_DATA = BASE_OUTPUT / "專員系統個資"
    PATH_ORDER = BASE_OUTPUT / "訂單資料"
    PATH_VIP = BASE_OUTPUT / "VIP儲值金"
    PATH_HR = BASE_OUTPUT / "服務分潤表"

else:
    BASE_GOOGLE_DRIVE = Path("/Users/jenny/Library/CloudStorage/GoogleDrive-jenny@lemonclean.com.tw/我的雲端硬碟")

    # Jenny 個人輸出
    PATH_JENNY = BASE_GOOGLE_DRIVE / "lemon_Jenny" / "Jenny@lemon程式"

    # 排班統計表
    PATH_SCHEDULE = Path("/Users/jenny/Library/CloudStorage/GoogleDrive-jenny@lemonclean.com.tw/.shortcut-targets-by-id/1zbu45AG1adMzz24HPdi_tLfh2Tncw_Br/排班統計表")

    # 其他輸出資料夾
    PATH_CLEANER_SCHEDULE = PATH_JENNY / "專員班表"
    PATH_CLEANER_DATA = PATH_JENNY / "專員系統個資"
    PATH_ORDER = PATH_JENNY / "訂單資料"

    # 財務
    PATH_VIP = BASE_GOOGLE_DRIVE / "lemon_財務" / "02.VIP儲值金"

    # 人事
    PATH_HR = BASE_GOOGLE_DRIVE / "lemon_人事" / "03 服務分潤表"

# 自動建立資料夾
for p in [
    PATH_JENNY,
    PATH_SCHEDULE,
    PATH_CLEANER_SCHEDULE,
    PATH_CLEANER_DATA,
    PATH_ORDER,
    PATH_VIP,
    PATH_HR,
]:
    p.mkdir(parents=True, exist_ok=True)

# 共用設定
API_LIMIT = 10000
