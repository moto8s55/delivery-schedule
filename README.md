# Delivery Schedule — Auto Generator

Calculates the delivery date from the latest order date in `orders.xlsx`  
(+2 business days, excluding weekends, Japanese public holidays, and custom holidays),  
then generates a single-line retro CRT HTML hosted on GitHub Pages for embedding in Notion.

---

## File Structure

```
delivery-schedule/
├── delivery.py                  # Main script
├── holidays.txt                 # Custom company holidays
├── orders.xlsx                  # Order data (auto-written by external system)
├── requirements.txt             # jpholiday, openpyxl
├── delivery.html                # Auto-generated HTML (do not edit manually)
└── .github/
    └── workflows/
        └── schedule.yml         # GitHub Actions: weekdays 17:00 JST only
```

---

## Setup

### 1. Create GitHub repository & push files

```bash
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/<your-username>/delivery-schedule.git
git push -u origin main
```

### 2. Enable GitHub Pages

1. Go to repository **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / `/ (root)` → Save

After a few minutes, the page will be live at:  
`https://<your-username>.github.io/delivery-schedule/delivery.html`

### 3. Embed in Notion

1. In your Notion page, type `/embed`
2. Paste the GitHub Pages URL above
3. Adjust embed height to ~50px

---

## orders.xlsx Format

| Column A (Order Date) | Column B+ (any) |
|-----------------------|-----------------|
| 2026/05/21            | ...             |
| 2026/05/26            | ...             |

- Row 1 is treated as a header (skipped)
- The **latest date** in column A is used for delivery calculation
- Date format: `YYYY/MM/DD` or `YYYY-MM-DD`
- This file is written automatically by your external system

---

## holidays.txt Format

```
# Year-end holidays
2026-12-29
2026-12-30
2026-12-31
```

- Lines starting with `#` are comments
- Weekends and Japanese public holidays are excluded automatically — no need to list them here
- Only add company-specific closures (inventory days, summer breaks, etc.)

---

## Changing the Skip Days

Edit the constant at the top of `delivery.py`:

```python
SKIP_DAYS = 2  # Change to 3 for 3 business days, etc.
```

---

## Auto-Update Schedule

GitHub Actions runs **every weekday at 17:00 JST** (UTC 08:00, Mon–Fri).  
Pushing `orders.xlsx` or `holidays.txt` also triggers an immediate rebuild.  
Weekends and public holidays: no execution.
