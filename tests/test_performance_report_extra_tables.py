import pandas as pd

import performance_report as report


def test_configurable_month_ranges_cross_year():
    assert report.get_report_month_ranges("2026-11", "2027-02") == [
        ("2026/11", "2026-11-01", "2026-11-30"),
        ("2026/12", "2026-12-01", "2026-12-31"),
        ("2027/01", "2027-01-01", "2027-01-31"),
        ("2027/02", "2027-02-01", "2027-02-28"),
    ]


def test_order_date_summary_groups_by_city_and_totals():
    # 跟 build_month_performance_summary 共用同一種 raw_df 形狀（來自同一個報表頁面的
    # parse_html() 結果），只是這裡不分月份。
    raw_df = pd.DataFrame([
        {"城市": "台北", "收入類型": "現金收入", "服務": "居家清潔", "已付款": 2000, "待付款": 1000},
        {"城市": "台中", "收入類型": "現金收入", "服務": "居家清潔", "已付款": 0, "待付款": 500},
    ])
    out = report.build_order_date_summary(raw_df)
    assert out.iloc[0].to_dict() == {
        "地區": "台北", "未付款": 1000, "已付款": 2000, "未付款＋已付款": 3000,
        "儲值金待付款": 0, "儲值金已付款": 0, "儲值金待付款＋已付款": 0,
    }
    total_row = out[out["地區"] == "加總"].iloc[0]
    assert total_row.to_dict() == {
        "地區": "加總", "未付款": 1500, "已付款": 2000, "未付款＋已付款": 3500,
        "儲值金待付款": 0, "儲值金已付款": 0, "儲值金待付款＋已付款": 0,
    }
    # 每個地區都要有自己的一列，不能像舊版那樣另外多一列「儲值金」。
    assert out["地區"].tolist() == [*report.CITY_ORDER, "加總"]


def test_order_date_summary_splits_out_stored_value_by_city():
    raw_df = pd.DataFrame([
        # 一般清潔訂單，即使是用儲值金付款，服務分類仍然是「清潔」，要留在未付款/已付款。
        {"城市": "台北", "收入類型": "現金收入", "服務": "居家清潔", "已付款": 2800, "待付款": 0},
        {"城市": "台中", "收入類型": "現金收入", "服務": "居家清潔", "已付款": 0, "待付款": 500},
        # 儲值金儲值單：raw_df 裡的「服務」欄位已經是 parse_html() 正規化過的值
        # （原始表頭「VIP」會被 normalize_service 轉成「儲值金」），收入類型是「現金收入」
        # ——跟「目前總表」判斷儲值金的邏輯完全相同。同一個地區（台北）同時有清潔訂單
        # 跟儲值金訂單，確認兩者不會互相汙染。
        {"城市": "台北", "收入類型": "現金收入", "服務": "儲值金", "已付款": 0, "待付款": 50000},
        {"城市": "桃園", "收入類型": "現金收入", "服務": "儲值金", "已付款": 30000, "待付款": 0},
    ])
    out = report.build_order_date_summary(raw_df)

    taipei_row = out[out["地區"] == "台北"].iloc[0]
    assert taipei_row.to_dict() == {
        "地區": "台北", "未付款": 0, "已付款": 2800, "未付款＋已付款": 2800,
        "儲值金待付款": 50000, "儲值金已付款": 0, "儲值金待付款＋已付款": 50000,
    }

    taoyuan_row = out[out["地區"] == "桃園"].iloc[0]
    assert taoyuan_row.to_dict() == {
        "地區": "桃園", "未付款": 0, "已付款": 0, "未付款＋已付款": 0,
        "儲值金待付款": 0, "儲值金已付款": 30000, "儲值金待付款＋已付款": 30000,
    }

    total_row = out[out["地區"] == "加總"].iloc[0]
    assert total_row.to_dict() == {
        "地區": "加總", "未付款": 500, "已付款": 2800, "未付款＋已付款": 3300,
        "儲值金待付款": 50000, "儲值金已付款": 30000, "儲值金待付款＋已付款": 80000,
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
