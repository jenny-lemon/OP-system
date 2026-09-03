from pathlib import Path

# 對齊 tool-system：儲值金表的週末加價不包含在付款金額內，
# 必須依本次查詢的付款狀態，加回待付款或已付款。
p = Path('performance_report.py')
s = p.read_text(encoding='utf-8')
old = '''            if income_type == "儲值金" and weekly_idx is not None and len(row) > weekly_idx:
                paid += safe_int(row[weekly_idx])
'''
new = '''            if income_type == "儲值金" and weekly_idx is not None and len(row) > weekly_idx:
                weekend_amount = safe_int(row[weekly_idx])
                if str(payment_status) == "0":
                    unpaid += weekend_amount
                else:
                    paid += weekend_amount
'''
if old not in s: raise SystemExit('weekend surcharge target not found')
s = s.replace(old, new, 1)
old = 'def parse_html(html):'
new = 'def parse_html(html, payment_status=None):'
if old not in s: raise SystemExit('parse_html signature target not found')
s = s.replace(old, new, 1)
start = s.index('def generate_order_date_report(')
end = s.index('\ndef generate_month_range_reports(', start)
block = s[start:end]
if 'parse_html(response.text)' not in block: raise SystemExit('order-date parse_html call not found')
block = block.replace('parse_html(response.text)', 'parse_html(response.text, payment_status=status)')
s = s[:start] + block + s[end:]
p.write_text(s, encoding='utf-8')

p = Path('order_date_service_report.py')
s = p.read_text(encoding='utf-8')
old = 'for item in report.parse_html(response.text):'
new = 'for item in report.parse_html(response.text, payment_status=status):'
if old not in s: raise SystemExit('service report parse_html target not found')
p.write_text(s.replace(old, new, 1), encoding='utf-8')

p = Path('tests/test_performance_report_extra_tables.py')
s = p.read_text(encoding='utf-8')
marker = '\ndef test_purchase_is_stored_value_topup_detects_order_title():'
test = '''\ndef test_parse_html_adds_stored_weekend_surcharge_to_matching_payment_status():
    html = """
    <table><tr><th>儲值金</th><th>已付款金額</th><th>待付款金額</th><th>週末加價</th></tr>
    <tr><td>居家清潔</td><td>1,000</td><td>2,000</td><td>200</td></tr></table>
    <table><tr><th>現金收入</th><th>已付款金額</th><th>待付款金額</th><th>週末加價</th></tr>
    <tr><td>居家清潔</td><td>3,000</td><td>4,000</td><td>300</td></tr></table>
    """
    paid_rows = report.parse_html(html, payment_status=1)
    unpaid_rows = report.parse_html(html, payment_status=0)
    assert paid_rows[0]["已付款"] == 1200
    assert paid_rows[0]["待付款"] == 2000
    assert unpaid_rows[0]["已付款"] == 1000
    assert unpaid_rows[0]["待付款"] == 2200
    assert paid_rows[1]["已付款"] == 3000
    assert unpaid_rows[1]["待付款"] == 4000

'''
if 'test_parse_html_adds_stored_weekend_surcharge_to_matching_payment_status' not in s:
    if marker not in s: raise SystemExit('test insertion marker not found')
    s = s.replace(marker, test + marker, 1)
p.write_text(s, encoding='utf-8')

# trigger corrected PYTHONPATH run
