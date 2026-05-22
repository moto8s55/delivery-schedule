"""
delivery.py - 搬入日自動計算 & HTML生成
- orders.xlsx から最新の注文日を取得
- 中2日（土日・祝日・独自休業日除く）で搬入日を計算
- レトロCRTデザイン1行HTMLを生成
"""

import datetime
import os

try:
    import jpholiday
except ImportError:
    jpholiday = None
    print("WARNING: jpholiday not installed. Holidays will not be excluded.")

try:
    import openpyxl
except ImportError:
    openpyxl = None
    print("WARNING: openpyxl not installed. Cannot read orders.xlsx.")

# ============================================================
# 設定
# ============================================================
HOLIDAYS_FILE = "holidays.txt"
ORDERS_FILE   = "orders.xlsx"
OUTPUT_HTML   = "delivery.html"
SKIP_DAYS     = 2  # 中N日

WEEKDAY_EN = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
MONTH_EN   = ["JAN","FEB","MAR","APR","MAY","JUN",
               "JUL","AUG","SEP","OCT","NOV","DEC"]


# ============================================================
# 独自休業日読み込み
# ============================================================
def load_custom_holidays(filepath: str) -> set:
    holidays = set()
    if not os.path.exists(filepath):
        return holidays
    with open(filepath, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            try:
                holidays.add(datetime.date.fromisoformat(line))
            except ValueError:
                print(f"  SKIP (format error): {line}")
    return holidays


# ============================================================
# 営業日判定
# ============================================================
def is_business_day(d: datetime.date, custom_holidays: set) -> bool:
    if d.weekday() >= 5:
        return False
    if d in custom_holidays:
        return False
    if jpholiday and jpholiday.is_holiday(d):
        return False
    return True


# ============================================================
# 搬入日計算（中N日後の営業日）
# ============================================================
def calc_delivery_date(order_date: datetime.date, custom_holidays: set, skip: int = 2) -> datetime.date:
    d = order_date
    count = 0
    while count < skip:
        d += datetime.timedelta(days=1)
        if is_business_day(d, custom_holidays):
            count += 1
    d += datetime.timedelta(days=1)
    while not is_business_day(d, custom_holidays):
        d += datetime.timedelta(days=1)
    return d


# ============================================================
# orders.xlsx から最新の注文日を取得
# ============================================================
def load_latest_order_date(filepath: str) -> datetime.date | None:
    if openpyxl is None:
        print("  ERROR: openpyxl not available.")
        return None
    if not os.path.exists(filepath):
        print(f"  ERROR: {filepath} not found.")
        return None

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    latest = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        raw = row[0]
        if isinstance(raw, datetime.datetime):
            d = raw.date()
        elif isinstance(raw, datetime.date):
            d = raw
        else:
            try:
                d = datetime.date.fromisoformat(str(raw).strip())
            except ValueError:
                continue
        if latest is None or d > latest:
            latest = d
    return latest


# ============================================================
# HTML生成（CRT 1行表示）
# ============================================================
def generate_html(delivery: datetime.date, now: datetime.datetime) -> str:
    wd  = WEEKDAY_EN[delivery.weekday()]
    mon = MONTH_EN[delivery.month - 1]
    delivery_str = f"{wd}, {mon} {delivery.day}"
    time_str     = now.strftime("%H:%M")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="refresh" content="3600">
  <title>Delivery Schedule</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap');
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      background: #020c02;
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
      padding: 12px;
    }}
    .crt {{
      font-family: 'Share Tech Mono', 'Courier New', monospace;
      background: #020c02;
      border: 1px solid #005520;
      border-radius: 6px;
      overflow: hidden;
      position: relative;
      box-shadow: 0 0 18px rgba(0,255,65,0.06);
      width: 100%;
      max-width: 480px;
    }}
    /* スキャンライン */
    .crt::before {{
      content: '';
      position: absolute; inset: 0;
      background: repeating-linear-gradient(
        0deg, transparent, transparent 2px,
        rgba(0,0,0,0.08) 2px, rgba(0,0,0,0.08) 4px
      );
      pointer-events: none;
      z-index: 10;
    }}
    .row {{
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 13px 20px;
    }}
    .lbl {{
      font-size: 12px;
      color: #007a1f;
      letter-spacing: 1px;
    }}
    .val {{
      font-size: 15px;
      color: #00ff41;
      font-weight: bold;
      text-shadow: 0 0 6px rgba(0,255,65,0.4);
    }}
    .sep {{
      font-size: 13px;
      color: #005510;
    }}
    @keyframes blink {{ 0%,100%{{opacity:1}} 50%{{opacity:0}} }}
    .cursor {{
      display: inline-block;
      width: 8px; height: 14px;
      background: #00ff41;
      vertical-align: middle;
      margin-left: 3px;
      animation: blink 1s step-end infinite;
    }}
  </style>
</head>
<body>
  <div class="crt">
    <div class="row">
      <span class="lbl">TIME</span>
      <span class="val">{time_str}</span>
      <span class="sep">&#9658;</span>
      <span class="lbl">DELIVERY</span>
      <span class="val">{delivery_str}<span class="cursor"></span></span>
    </div>
  </div>
</body>
</html>
"""


# ============================================================
# メイン
# ============================================================
def main():
    now = datetime.datetime.now()
    print(f"[{now:%Y/%m/%d %H:%M:%S}] delivery.py started")

    # 独自休業日読み込み
    custom_holidays = load_custom_holidays(HOLIDAYS_FILE)
    print(f"  Custom holidays: {len(custom_holidays)} days loaded")

    # 注文日取得
    order_date = load_latest_order_date(ORDERS_FILE)
    if order_date is None:
        # フォールバック: 今日を注文日とする
        order_date = now.date()
        print(f"  Fallback: using today as order date ({order_date})")
    else:
        print(f"  Order date: {order_date}")

    # 搬入日計算
    delivery = calc_delivery_date(order_date, custom_holidays, SKIP_DAYS)
    wd  = ["MON","TUE","WED","THU","FRI","SAT","SUN"][delivery.weekday()]
    mon = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][delivery.month-1]
    print(f"  Delivery date: {delivery} ({wd}, {mon} {delivery.day})")

    # HTML生成
    html = generate_html(delivery, now)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Generated: {OUTPUT_HTML}")
    print("done.")


if __name__ == "__main__":
    main()
