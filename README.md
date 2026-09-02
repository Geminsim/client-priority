# Lead Priority Excel Add-in

A local Microsoft Excel Office Add-in that scores prospective customers, assigns A–D priorities, sorts the lead list, and creates a dashboard with summary metrics and charts. All data stays inside the workbook.

## Features

- Detects common English and Chinese lead-data headers.
- Calculates a normalized 0–100 priority score using adjustable weights.
- Assigns clear A, B, C, or D follow-up priorities.
- Adds a concise scoring rationale for every lead.
- Sorts leads by priority score.
- Creates a Lead Priority Dashboard with distribution and weight charts.

## Data format

The first row must contain headers, including at least a `Company` column. Recommended columns accept values on a 0–100, 1–10, or 1–5 scale:

| Company | Customer Fit | Buying Intent | Budget | Authority | Urgency |
|---|---:|---:|---:|---:|---:|
| Northstar Technology | 90 | 80 | 70 | 60 | 50 |

Missing scoring columns are treated as zero. Common alternative headers such as `Account`, `Lead`, `Fit`, and localized Chinese equivalents are recognized automatically.

## Run locally on Windows with desktop Excel

1. Install Node.js 18 or later.
2. Run `npm install` in this directory.
3. Run `npm start`. The first run installs a local HTTPS certificate and sideloads `manifest.xml` into Excel.
4. In Excel, open **Home → Lead Analysis → Priority Assistant**.
5. Run `npm stop` when you finish debugging.

If your organization blocks automatic sideloading, run `npm run serve`, then use Excel's add-in management screen to upload `manifest.xml` manually.

## Scoring model

Default weights are Customer Fit 30%, Buying Intent 25%, Budget 20%, Decision Authority 15%, and Urgency 10%. The add-in normalizes the configured weights automatically.

- A: score of 80 or higher — follow up now
- B: 60–79 — nurture actively
- C: 40–59 — monitor
- D: below 40 — deprioritize

Each analysis adds or updates `Priority Score`, `Lead Priority`, and `Scoring Rationale` columns.

## Test

Run `npm test` to test the scoring engine without starting Excel.

## Privacy

The add-in does not send workbook data to a server. Processing happens through the Excel JavaScript API in the open workbook.
