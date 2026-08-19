import pandas as pd

import performance_report as report


def test_configurable_month_ranges_cross_year():
    assert report.get_report_month_ranges("2026-11", "2027-02") == [
        ("2026/11", "2026-11-01", "2026-11-30"),
        ("2026/12", "2026-12-01", "2026-12-31"),
        ("2027/01", "2027-01-01", "2027-01-31"),
        ("2027/02", "2027-02-01", "2027-02-28"),
    ]


def test_order_date_summary_excludes_cancelled_orders():
    records = [
        {"__city": "台北", "total": "1,000", "purchase_status": "0"},
        {"__city": "台北", "total": "2,000", "purchase_status": "1"},
        {"__city": "台中", "total": 500, "purchase_status": "0"},
        {"__city": "台北", "total": 999, "purchase_status": "1", "cancel_at": "2026-08-18"},
    ]
    out = report.build_order_date_summary(records)
    assert out.iloc[0].to_dict() == {
        "地區": "台北", "未付款": 1000, "已付款": 2000, "未付款＋已付款": 3000,
    }
    assert out.iloc[-1].to_dict() == {
        "地區": "加總", "未付款": 1500, "已付款": 2000, "未付款＋已付款": 3500,
    }


def test_order_date_summary_splits_out_stored_value_topups():
    records = [
        {"__city": "台北", "total": "1,000", "purchase_status": "0"},
        {"__city": "台北", "total": "2,000", "purchase_status": "1"},
        {"__city": "台中", "total": 500, "purchase_status": "0"},
        # 儲值金訂單本身：購買項目開頭是「儲值金」，待付款。
        {"__city": "台北", "total": 50000, "purchase_status": "0", "buy": "儲值金-台北(儲值金50,000贈購物金2,500)"},
        # 儲值金訂單本身：購買項目開頭是「儲值金」，已付款。
        {"__city": "桃園", "total": 30000, "purchase_status": "1", "buy": "儲值金-桃園(儲值金30,000)"},
        # 用儲值金「付款」的清潔訂單：付款方式是儲值金，但購買項目是居家清潔，不能被當成儲值單。
        {"__city": "台北", "total": 2800, "purchase_status": "1", "buy": "居家清潔", "payway": "儲值金"},
    ]
    out = report.build_order_date_summary(records)

    assert out.iloc[0].to_dict() == {
        "地區": "台北", "未付款": 1000, "已付款": 2000 + 2800, "未付款＋已付款": 1000 + 2000 + 2800,
    }

    total_row = out[out["地區"] == "加總"].iloc[0]
    assert total_row.to_dict() == {
        "地區": "加總", "未付款": 1500, "已付款": 4800, "未付款＋已付款": 6300,
    }

    stored_value_row = out.iloc[-1]
    assert stored_value_row.to_dict() == {
        "地區": "儲值金", "未付款": 50000, "已付款": 30000, "未付款＋已付款": 80000,
    }


def test_month_performance_minus_reserve_equals_net():
    month_ranges = [("2026/08", "2026-08-01", "2026-08-31")]
    raw_df = pd.DataFrame([
        {"城市": "台北", "月份": "2026/08", "收入類型": "現金收入", "服務": "居家清潔", "已付款": 8000, "待付款": 2000},
        {"城市": "台中", "月份": "2026/08", "收入類型": "現金收入", "服務": "居家清潔", "已付款": 8000, "待付款": 0},
    ])
    reserve_records = [
        {"__city": "台北", "order_no": "LC1", "date_clean": "2026-08-20", "name": "檸檬保留", "person": 2, "hour": 4, "total": 4800},
        {"__city": "台中", "order_no": "LC2", "date_clean": "2026-08-21", "notice": "大掃除檸檬保留單", "person": 2, "period_s": "09:00", "period_e": "12:00", "total": 3600},
    ]
    performance_df = report.build_month_performance_summary(raw_df, month_ranges)
    reserve_df = report.build_reserve_summary(reserve_records, month_ranges)
    net_df = report.build_net_performance_summary(raw_df, reserve_df, month_ranges)

    assert reserve_df[reserve_df["地區"] == "台北"].iloc[0]["2026/08保留單時數"] == 8
    assert reserve_df[reserve_df["地區"] == "台中"].iloc[0]["2026/08保留單時數"] == 6

    for city in [*report.CITY_ORDER, "加總"]:
        gross = performance_df[performance_df["地區"] == city].iloc[0]["2026/08業績"]
        reserve = reserve_df[reserve_df["地區"] == city].iloc[0]["2026/08保留單業績"]
        net = net_df[net_df["地區"] == city].iloc[0]["2026/08業績－保留單業績"]
        assert gross - reserve == net
