# Client Delivery Tracker

An Excel workbook for logging client requests, tracking them against service-level targets, and producing a weekly report. It is built for the day-to-day work of an IT-services trainee or coordinator: keep requests organised, spot what is overdue, and report clearly to a manager.

All data in this repository is **sample data** (fictional clients) for demonstration.

## What it does

| Sheet | Purpose |
|:--|:--|
| **Requests** | Log of requests as an Excel Table. Dropdowns for type, priority and status; formulas for days open, SLA target and SLA status; overdue and missed items highlighted red, met items green. |
| **Dashboard** | KPIs (total, still open, overdue, % SLA met, average days to resolve), counts by status, type and client, and a status chart, all formula-driven. |
| **Weekly report** | Sentences generated from the dashboard, plus a list of overdue items, ready to paste into an email or Word. |
| **Settings** | SLA target days per priority and the report date. Change them and everything recalculates. |

Excel features used: tables, data validation, `VLOOKUP`, `IF`, `COUNTIF(S)`, `AVERAGEIF`, `TEXTJOIN`, conditional formatting, charts, freeze panes.

## Use it

1. Build the workbook: `pip install openpyxl && python build_tracker.py` (reads `sample_requests.csv`, writes `Client_Delivery_Tracker.xlsx`).
2. Add rows on **Requests**; set the report date on **Settings**.
3. Read the **Dashboard**; copy the **Weekly report** text into your update.

## Notes

- Built with the help of an AI assistant (Claude); I reviewed the formulas and can explain each one.
- Closed or resolved requests count as met or missed against the SLA; open ones count as on track or overdue.

Author: Jayanta Roy
