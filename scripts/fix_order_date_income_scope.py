from pathlib import Path

p = Path('performance_report.py')
s = p.read_text(encoding='utf-8')
old = 'service_df = work[work["類別"] != "儲值金"].copy()'
new = 'service_df = work[(work["收入類型"] == "現金收入") & (work["類別"] != "儲值金")].copy()'
if old not in s:
    raise SystemExit('service_df target not found')
s = s.replace(old, new, 1)
old = '''                        category = to_category(row["服務"], row["收入類型"])
                        if category == "儲值金":
                            continue
'''
new = '''                        category = to_category(row["服務"], row["收入類型"])
                        # 訂購日期付款主表要對應後台「現金收入」統計表；
                        # VIP/儲值金付款的服務不可再疊進一般待付款/已付款。
                        if row["收入類型"] != "現金收入" or category == "儲值金":
                            continue
'''
if old not in s:
    raise SystemExit('month totals target not found')
s = s.replace(old, new, 1)
p.write_text(s, encoding='utf-8')

p = Path('order_date_service_report.py')
s = p.read_text(encoding='utf-8')
old = '''                    for item in report.parse_html(response.text):
                        category = report.to_category(item["服務"], item["收入類型"])
'''
new = '''                    for item in report.parse_html(response.text):
                        # 家電/水洗付款表同樣對應後台「現金收入」統計表，
                        # 不把 VIP 或儲值金付款服務重複加進來。
                        if item["收入類型"] != "現金收入":
                            continue
                        category = report.to_category(item["服務"], item["收入類型"])
'''
if old not in s:
    raise SystemExit('service report target not found')
s = s.replace(old, new, 1)
p.write_text(s, encoding='utf-8')
