# AP / AR Tracker with Aging Report & Dashboard

An automated Accounts Payable and Accounts Receivable tracker built entirely in Microsoft Excel. Designed for finance teams and small business owners who want to replace manual spreadsheet tracking with a clean, formula-driven workbook that updates automatically as data is entered.

---

## What It Does

Managing invoices and vendor bills manually is slow and error-prone. This workbook solves that by:

- Automatically calculating how overdue each invoice is
- Grouping outstanding balances into aging buckets (Current, 1–30, 31–60, 61–90, 90+ days)
- Showing a live dashboard with key financial metrics and charts
- Generating a printable aging report with zero manual calculation

---

## Features

### AR Ledger (Accounts Receivable)
Track every invoice you have sent to clients.
- Enter the invoice number, date, client, and amount
- The workbook auto-calculates tax, total, outstanding balance, days overdue, aging bucket, and status
- Status updates automatically: **Paid**, **Partial**, **Outstanding**, or **Overdue**

### AP Ledger (Accounts Payable)
Track every bill you owe to vendors.
- Same structure as AR Ledger
- Includes a **Next Payment Due** column on the aging report so you never miss a payment

### Aging Reports
One summary page each for AR and AP.
- One row per client or vendor
- Columns broken down by aging bucket: Current → 1–30 → 31–60 → 61–90 → 90+ days
- Totals row at the bottom
- Colour-coded: 🔴 red for 90+ days, 🟠 amber for 61–90, 🟡 yellow for 31–60, 🟢 green for current/paid
- Print-ready on A4 portrait

### Dashboard
A one-page summary for management reporting.
- **6 KPI cards**: Total AR Outstanding, Total AP Outstanding, Net Position, Total 90+ Overdue, Number of Overdue Invoices, Collection Rate %
- **AR Aging Distribution chart** — horizontal bar chart showing balance per aging bucket
- **AP Aging Distribution chart** — same for vendor payables
- All values update automatically as you enter data in the ledgers

### Settings
Customise the workbook for your business without touching any formulas.
- Company name (appears in report headers)
- Tax rate (default: 6% SST)
- Aging thresholds (default: 30 / 60 / 90 days)
- Currency symbol
- Invoice and bill number prefixes

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| Dashboard | Live KPIs and charts — no data entry here |
| AR Ledger | Client invoice records |
| AP Ledger | Vendor bill records |
| AR Aging Report | Aging breakdown by client |
| AP Aging Report | Aging breakdown by vendor |
| Settings | Company config and thresholds |
| Lookup Tables | Client list, vendor list, payment terms, account codes |

---

## How to Use

### First-time setup
1. Open `AP_AR_Tracker.xlsx` in Microsoft Excel
2. Go to the **Settings** sheet and update:
   - Company Name
   - Tax Rate (if not 6%)
   - Invoice and bill prefixes (e.g. `INV-`, `BILL-`)
3. Go to **Lookup Tables** and replace the sample client and vendor names with your real ones

### Entering invoices (AR)
1. Go to the **AR Ledger** sheet
2. Fill in the **yellow cells** only:
   - Invoice Number, Invoice Date, Due Date
   - Client Name (select from dropdown)
   - Description, Amount (excl. Tax)
   - When payment is received: Payment Date and Payment Amount
3. The **grey cells** (Tax, Total, Outstanding Balance, Days Overdue, Aging Bucket, Status) calculate automatically — do not type in them

### Entering bills (AP)
1. Go to the **AP Ledger** sheet
2. Same process as AR Ledger — fill in the yellow cells

### Adding new rows
- Click on the last row of the table and press **Tab** to add a new row
- All formulas carry over automatically

### Viewing your aging report
- Navigate to **AR Aging Report** or **AP Aging Report**
- Balances update in real time as you enter data in the ledgers
- To print: File → Print (already configured for A4 portrait)

### Updating the dashboard
- The Dashboard updates automatically — no action needed
- If values appear stale after a large data entry, press **Ctrl + Alt + F9** to force a full recalculate

---

## Requirements

- Microsoft Excel 2019 or later (Excel 365 recommended)
- No macros, no plugins, no internet connection required
- Works on Windows and Mac

---

## Rebuilding from Scratch (Optional)

A Python script is included to regenerate the entire workbook from scratch. This is useful if you want to reset to the sample data or redeploy the template.

**Requirements:**
```
Python 3.x
pip install openpyxl
```

**Run:**
```bash
python build_workbook.py
```

This will regenerate `AP_AR_Tracker.xlsx` with the original sample data and all formatting intact.

---

## Currency & Localisation

- Default currency: Malaysian Ringgit (RM)
- Default tax: SST 6%
- Date format: DD/MM/YYYY
- All of the above are configurable in the **Settings** sheet

---

## Roadmap (Phase 2)

- Multi-currency support with live exchange rates via Power Query
- Automated overdue reminder email generator (Python integration)
- CSV bank statement import
- Power BI external dashboard connection
- Password protection for formula sheets

---

## Author

**Zahrin Bin Jasni**
Built for practical finance operations and portfolio demonstration.
