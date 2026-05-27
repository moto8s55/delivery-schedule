"""
delivery.py - 搬入日HTML生成
- holidays.txt から休業日リストを読み込む
- 搬入日の計算はブラウザ側JS（毎秒リアルタイム更新）
- GitHub Actions（Ubuntu）で動作するよう Excelへの依存を除去
"""

import datetime
from pathlib import Path

OUTPUT_HTML   = Path(__file__).parent / "delivery.html"
HOLIDAYS_FILE = Path(__file__).parent / "holidays.txt"


# ============================================================
# holidays.txt を読み込む
# 書式：1行1日付、YYYY-MM-DD、# でコメント
# ============================================================
def load_holidays(path: Path) -> list[str]:
    if not path.exists():
        print(f"  WARN: {path.name} not found. No custom holidays.")
        return []
    holidays = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#")[0].strip()
        if line:
            holidays.append(line)
    print(f"  Holidays loaded: {len(holidays)} days")
    return holidays


# ============================================================
# HTML生成
# 搬入日ロジックはすべてJS側で実装（毎秒更新）
# ============================================================
def generate_html(holidays: list[str]) -> str:
    # JS配列リテラルを生成
    holidays_js = ",\n      ".join(f'"{h}"' for h in sorted(holidays))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
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
      <span class="val" id="clock">--:--</span>
      <span class="sep">&#9658;</span>
      <span class="lbl">DELIVERY</span>
      <span class="val" id="delivery">--<span class="cursor"></span></span>
    </div>
  </div>
  <script>
    // ------------------------------------------------
    // holidays.txt の内容をPythonが埋め込む
    // ------------------------------------------------
    const HOLIDAYS = new Set([
      {holidays_js}
    ]);

    // ------------------------------------------------
    // ユーティリティ
    // ------------------------------------------------
    function toStr(d) {{
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + dd;
    }}

    function addDays(d, n) {{
      const r = new Date(d);
      r.setDate(r.getDate() + n);
      return r;
    }}

    function isWorkday(d) {{
      const w = d.getDay();
      if (w === 0 || w === 6) return false;   // 土日
      if (HOLIDAYS.has(toStr(d))) return false; // 祝日・休業日
      return true;
    }}

    // ------------------------------------------------
    // 搬入日計算
    //
    // ルール：
    //   月〜水 17時前  → 今週木曜
    //   水    17時以降 → 翌週月曜
    //   木〜金 17時前  → 翌週月曜
    //   木〜金 17時以降→ 翌週火曜方向（次の月曜の翌営業日）
    //   土・日         → 翌週月曜方向
    //
    //   ※ 搬入候補日（木 or 月）が祝日の場合は次の営業日にスライド
    // ------------------------------------------------
    function getNextDelivery(now) {{
      const day  = now.getDay();   // 0=日,1=月,2=火,3=水,4=木,5=金,6=土
      const hour = now.getHours(); // ローカル時刻（JST）
      const CUTOFF = 17;

      // 17時以降は「翌日」を起点にして計算
      let base = new Date(now);
      base.setHours(0, 0, 0, 0);
      if (hour >= CUTOFF) {{
        base = addDays(base, 1);
      }}
      const baseDay = base.getDay();

      // 次の「木曜」または「月曜」を目標曜日として選ぶ
      let targetDay;
      if (baseDay === 1 || baseDay === 2 || baseDay === 3) {{
        // 月・火・水（17時前）→ 今週木曜
        targetDay = 4;
      }} else {{
        // 水17時以降(base=木), 木17時前, 木17時以降(base=金),
        // 金, 土, 日 → 翌週月曜
        targetDay = 1;
      }}

      let diff = targetDay - baseDay;
      if (diff <= 0) diff += 7;

      let delivery = addDays(base, diff);

      // 祝日・休業日ならスライド
      while (!isWorkday(delivery)) {{
        delivery = addDays(delivery, 1);
      }}

      return delivery;
    }}

    // ------------------------------------------------
    // 表示更新（毎秒）
    // ------------------------------------------------
    const DAYS   = ['SUN','MON','TUE','WED','THU','FRI','SAT'];
    const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN',
                    'JUL','AUG','SEP','OCT','NOV','DEC'];

    function tick() {{
      const now = new Date();

      // 時計
      const hh = String(now.getHours()).padStart(2, '0');
      const mm = String(now.getMinutes()).padStart(2, '0');
      document.getElementById('clock').textContent = hh + ':' + mm;

      // 搬入日（カーソルspanを残しつつテキストだけ更新）
      const d = getNextDelivery(now);
      const deliveryEl = document.getElementById('delivery');
      deliveryEl.firstChild.textContent =
        DAYS[d.getDay()] + ', ' + MONTHS[d.getMonth()] + ' ' + d.getDate();
    }}

    tick();
    setInterval(tick, 1000);
  </script>
</body>
</html>
"""


# ============================================================
# メイン
# ============================================================
def main():
    now = datetime.datetime.now()
    print(f"[{now:%Y/%m/%d %H:%M:%S}] delivery.py started")

    holidays = load_holidays(HOLIDAYS_FILE)
    html = generate_html(holidays)

    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"  Generated: {OUTPUT_HTML}")
    print("done.")


if __name__ == "__main__":
    main()
