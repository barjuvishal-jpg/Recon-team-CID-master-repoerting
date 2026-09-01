# Recon Team CID Master Reporting Dashboard & WebApp

A comprehensive Google Apps Script & Google Sheets real-time operational analytics dashboard and reporting tool for the Review Team under S.T.E.P.S Global Solutions.

## 🚀 Key Features

- **📊 Dashboard Summary:**
  - Real-time KPIs (Employees, Activities, Top Performer, Total Master Rows).
  - ⚠️ Unknown Activities Inspection Modal (detects missing dates or employee names with detailed diagnostics).
  - Date Range Filtering (Date From / Date To).
  - Top 10 Performers with workload volume and percentage progress bars.
  - Multi-sheet CSV and Excel export.

- **🏆 Performance Leaderboard:**
  - Full employee ranking table with Gold, Silver, and Bronze badges.
  - Total activity counts and percentage contribution share.
  - Direct click-to-profile navigation.

- **👥 Employees Profile Drill-Down:**
  - Individual employee summary card with rank badge and total activity counter.
  - Emerald Green Daily Activity Trend SVG Line Chart with area gradient and tooltips.
  - Category and section activity breakdown table.
  - SVG Multi-Slice Donut Chart with color-coded legend.
  - Individual profile CSV report export.

## 📁 Repository Structure

- `Dashboard.html` - Complete responsive web interface (HTML5, CSS3, Vanilla JS, SVG Charts).
- `WebApp.js` - Single-pass Apps Script backend aggregation engine, metrics calculator, and export handlers.
- `Code.js` - Core Google Sheets sync and background automation script.
- `appsscript.json` - Google Apps Script project manifest.
- `.clasp.json` - Clasp project configuration.
