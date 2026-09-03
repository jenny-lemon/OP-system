from pathlib import Path

# 訂購日期付款彙總仍要把「現金收入＋儲值金付款服務」一起加總；
# 問題是 parse_html() 會把儲值金表的「週末加價」另外加回已付款，
# 造成訂購日期付款彙總比後台「已付款金額」欄多算一次週末加價。
p = Path('performance_report.py')
s = p.read_text(encoding='utf-8')

old = 'def parse_html(html):'
new = 'def parse_html(html, include_stored_value_weekend_surcharge=True):'
if old not in s:
    raise SystemExit('parse_html signature target not found')
s = s.replace(old, new, 1)

old = '            if income_type == "儲值金" and weekly_idx is not None and len(row) > weekly_idx:\n                paid += safe_int(row[weekly_idx])'
new = '            if include_stored_value_weekend_surcharge and income_type == "儲值金" and weekly_idx is not None and len(row) > weekly_idx:\n                paid += safe_int(row[weekly_idx])'
if old not in s:
    raise SystemExit('weekly surcharge target not found')
s = s.replace(old, new, 1)

# 只在 generate_order_date_report() 範圍內關掉週末加價加回；
# 其他目前總表／月份報表維持既有口徑，避免副作用。
start = s.index('def generate_order_date_report(')
end = s.index('\ndef generate_month_range_reports(', start)
block = s[start:end]
old_call = 'parse_html(response.text)'
count = block.count(old_call)
if count != 2:
    raise SystemExit(f'expected 2 order-date parse_html calls, got {count}')
block = block.replace(
    old_call,
    'parse_html(response.text, include_stored_value_weekend_surcharge=False)',
)
s = s[:start] + block + s[end:]
p.write_text(s, encoding='utf-8')

# 家電／水洗是訂購日期付款彙總的延伸，也要使用相同口徑：
# 現金收入＋儲值金付款服務都保留，但不額外把週末加價塞進「已付款」。
p = Path('order_date_service_report.py')
s = p.read_text(encoding='utf-8')
old = 'for item in report.parse_html(response.text):'
new = 'for item in report.parse_html(response.text, include_stored_value_weekend_surcharge=False):'
if old not in s:
    raise SystemExit('service report parse_html target not found')
s = s.replace(old, new, 1)
p.write_text(s, encoding='utf-8')
