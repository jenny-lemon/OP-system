from auth import get_account, login

# 之後會引入報表模組
# 先用 print 模擬
def run_schedule_stats(session, city):
    print(f"{city} → 排班統計表")

def run_cleaner_schedule(session, city):
    print(f"{city} → 專員班表")

def run_cleaner_data(session, city):
    print(f"{city} → 專員個資")

def run_order_report(session, city):
    print(f"{city} → 訂單資料")


def run_city_bundle(city):
    print(f"\n=== 開始 {city} ===")

    acc = get_account(city)
    session = login(acc["email"], acc["password"])

    run_schedule_stats(session, city)
    run_cleaner_schedule(session, city)
    run_cleaner_data(session, city)
    run_order_report(session, city)

    print(f"=== 結束 {city} ===")


def main():
    for city in ["台北", "台中"]:
        run_city_bundle(city)


if __name__ == "__main__":
    main()
