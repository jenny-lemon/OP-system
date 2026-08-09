from pathlib import Path

HOME = Path.home()

LOCAL_GOOGLE_DRIVE = Path(
    "/Users/jenny/Library/CloudStorage/"
    "GoogleDrive-jenny@lemonclean.com.tw/我的雲端硬碟"
)

IS_LOCAL_MAC = (
    str(HOME).startswith("/Users/")
    and LOCAL_GOOGLE_DRIVE.exists()
)

CLOUD_BASE = Path("/tmp/lemon_data")

if IS_LOCAL_MAC:
    PATH_REPORT = (
        LOCAL_GOOGLE_DRIVE
        / "lemon_Jenny"
        / "Jenny@lemon程式"
        / "業績報表"
    )
else:
    PATH_REPORT = CLOUD_BASE / "業績報表"

try:
    PATH_REPORT.mkdir(parents=True, exist_ok=True)
except Exception:
    pass

API_LIMIT = 10000
