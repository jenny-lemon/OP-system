from fastapi import FastAPI, Header, HTTPException
from performance_report import generate_sales_report

app = FastAPI()

API_TOKEN = "換成你自己的長隨機字串"

@app.get("/trigger-performance-report")
def trigger_performance_report(x_api_token: str = Header(default="")):
    if x_api_token != API_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized")

    result = generate_sales_report(
        send_email=False,
        persist_dashboard=True,
        trigger="schedule"
    )

    return {
        "status": "ok",
        "updated_at": result.get("updated_at"),
        "error": result.get("error"),
        "df4_rows": len(result.get("df4", [])),
        "daily_rows": len(result.get("daily_df", [])),
    }
