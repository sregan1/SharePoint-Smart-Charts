# SharePoint Smart Charts

[![Version](https://img.shields.io/badge/version-1.3.1-informational)](CHANGELOG.md) [![Website](https://img.shields.io/badge/Website-sharepointsmartsolutions.com-blue)](https://sharepointsmartsolutions.com/sharepoint-smart-charts) [![User Guide](https://img.shields.io/badge/User%20Guide-Read%20Now-green)](USER-GUIDE.md) [![Download](https://img.shields.io/badge/Download-Latest%20Release-CA5010?logo=github&logoColor=white)](../../releases/latest) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A SharePoint Framework (SPFx) web part that renders interactive charts from multiple data sources with no coding required — 17 chart types, drill-down, bookmarks, trendlines, and click-to-filter integration, all configured through the SharePoint page editor.

![SPFx](https://img.shields.io/badge/SPFx-1.20-0078D4?logo=microsoft&logoColor=white) ![React](https://img.shields.io/badge/React-17-61DAFB?logo=react&logoColor=black) ![TypeScript](https://img.shields.io/badge/TypeScript-4.7-3178C6?logo=typescript&logoColor=white) ![Chart.js](https://img.shields.io/badge/Chart.js-4.x-FF6384?logo=chartdotjs&logoColor=white) ![PnPjs](https://img.shields.io/badge/PnPjs-3.22-0078D4)

![Hero chart with web part header](screenshots/hero-chart.png)

---

## Features

### Adaptive Property Pane

The property pane is organized into three focused pages — **Chart**, **Appearance**, and **Advanced** — with settings that adapt to the selected chart type. Options that don't apply to the current chart are hidden automatically, so the pane stays clear and relevant without overwhelming you. All 17 chart types are always accessible.

### Data Sources

| Source | Description |
|---|---|
| **Upload File** | Upload a CSV, TSV, or Excel (.xlsx / .xls) file directly from your computer. Data persists across page reloads (up to 200 KB). Multi-sheet workbooks get a sheet picker. |
| **SharePoint List** | Load data from any SharePoint list on the current or another site (up to 5,000 items, with a warning when truncated) |
| **SharePoint File** | Reference a CSV or Excel file stored in a SharePoint document library by URL |
| **REST API** | Connect to any REST endpoint that returns JSON, with optional response caching |
| **Microsoft Graph** | Chart any Graph endpoint (e.g. `/me/memberOf`) — requires admin-approved permissions |

### Chart Types

![Gallery showing chart types](screenshots/chart-gallery.png)

| Chart | Best Used For |
|---|---|
| **Bar (Vertical)** | Comparing values across categories |
| **Bar (Horizontal)** | Same as vertical; better for long category names |
| **Line** | Trends over time |
| **Area** | Trends over time with emphasis on volume |
| **Scatter** | Correlation between two numeric variables |
| **Bubble** | Correlation between two numeric variables with a third size dimension |
| **Pie** | Part-to-whole relationships (up to ~7 categories) |
| **Doughnut** | Same as pie with a center hole for labels or KPIs |
| **Radar** | Multi-attribute comparison across several items |
| **KPI Tile** | A single aggregated number with threshold coloring — dashboard headlines |
| **Histogram** | Distribution of a numeric column, with configurable bins |
| **Waterfall** | Cumulative gains and losses (e.g. monthly P&L) |
| **Box Plot** | Value distribution and outliers per category |
| **Violin Plot** | Distribution shape and density per category |
| **Before-After Plot** | Paired change from before to after values per row |
| **Treemap** | Proportions as nested area tiles |
| **Heatmap** | Intensity across two category dimensions |

### Data Shaping & Analytics

- **Group-by aggregation** — Sum, Average, Count, Min, or Max per category, applied before charting
- **Sort, filter, and row limit** — shape the data inline without touching the source
- **Trendlines & forecast** — linear regression or moving average overlays; project the trend up to 12 periods ahead
- **Reference lines** — fixed value, mean, or median drawn as a dashed line
- **Date/time X axis** — date columns are auto-detected and plotted on a true time scale
- **Logarithmic axes** — Y axis on most chart types, X axis on Scatter/Bubble, independently on a secondary axis
- **Conditional formatting** — highlight values above or below a threshold in a custom color
- **Combo charts** — mix bars and lines per series on the same chart
- **Dual Y axis** — plot a subset of series against a second, independently scaled right-hand axis (Bar, Line, Area)
- **Error bars** — standard deviation, standard error, or a custom per-row margin column (Bar, Line, Area)
- **Significance brackets** — annotate statistical comparisons between bar groups with a label, e.g. `*` or `p<0.05` (Bar)

### Interactivity

![Drill-down with breadcrumb navigation](screenshots/feature-drilldown.png)

- **Viewer filters** — opt-in filter bar so page visitors can slice the data themselves
- **Details on demand** — click a bar or slice to see the underlying rows
- **Drill-down** — define a column hierarchy (e.g. Region → Country → City) and click to drill, with breadcrumb navigation
- **Bookmarks** — save named view states; viewers pick them from a dropdown
- **Dynamic Data** — chart clicks publish the selected category/value/series to connected web parts
- **Custom tooltips** — append extra columns from the hovered row
- **Auto-refresh & caching** — reload network sources on an interval, cache API responses per session

### Other Features

- **Web part header** — optional title above the chart, toggled from the property pane
- **7 color palettes** — Office, Vibrant, Pastel, Monochrome, Traffic Light, Warm, Cool — plus per-series color overrides and color-by-category for scatter/bubble
- **Data labels** — optional value annotations with prefix/suffix and K/M abbreviation
- **Stacked bars** — toggle stacking on Bar and Line charts
- **Step interpolation** — render Line/Area charts with stepped segments instead of straight/smooth ones
- **Data point overlay** — show individual value markers on top of Bar/Horizontal Bar charts
- **Bubble size legend** — optional small/medium/large size key for the Bubble chart
- **Data table** — optional tabular view below the chart, paginated
- **Export** — download as PNG, JPEG, CSV (UTF-8 with BOM, Excel-safe), or Excel from every chart
- **Dark theme support** — chart colors adapt to dark section backgrounds
- **Localized & accessible** — all UI strings localizable; screen-reader-friendly charts and forms

---

## Screenshots

See [Chart Types](#chart-types) above for individual chart screenshots.

| | |
|---|---|
| ![Data source — empty state](screenshots/datasource-empty.png) | ![Data source — file loaded](screenshots/datasource-file-loaded.png) |
| **Data Source — Empty** | **Data Source — File Loaded** |
| ![Hero chart with web part header](screenshots/hero-chart.png) | ![SharePoint list data source](screenshots/datasource-sharepoint-list.png) |
| **Hero Chart** | **SharePoint List Source** |
| ![Data labels on bar chart](screenshots/feature-data-labels.png) | ![Color palette swatches](screenshots/feature-palettes.png) |
| **Data Labels** | **Color Palettes** |
| ![Data table below chart](screenshots/feature-data-table.png) | ![Stacked bar chart](screenshots/feature-stacked.png) |
| **Data Table** | **Stacked Mode** |
| ![Property pane — chart settings](screenshots/settings-chart-settings.png) | ![Property pane — data labels and axes](screenshots/settings-data-labels.png) |
| **Property Pane — Chart Settings** | **Property Pane — Data Labels & Axes** |
| ![KPI tile](screenshots/chart-kpi.png) | ![Histogram](screenshots/chart-histogram.png) |
| **KPI Tile** | **Histogram** |
| ![Waterfall chart](screenshots/chart-waterfall.png) | ![Box plot](screenshots/chart-boxplot.png) |
| **Waterfall** | **Box Plot** |
| ![Treemap](screenshots/chart-treemap.png) | ![Heatmap](screenshots/chart-heatmap.png) |
| **Treemap** | **Heatmap** |
| ![Viewer filter bar](screenshots/feature-viewer-filters.png) | ![Details on demand](screenshots/feature-details-on-demand.png) |
| **Viewer Filter Bar** | **Details on Demand** |
| ![Advanced options panel](screenshots/feature-advanced-options.png) | ![Property pane — advanced page](screenshots/settings-advanced.png) |
| **Advanced Options Panel** | **Property Pane — Advanced Page** |
| ![Microsoft Graph source with Data Path mapped to a JSON response](screenshots/graphapi-datapath.png) | ![SharePoint Admin Center — API access, Graph scopes pending approval](screenshots/admin-graph-consent.png) |
| **Microsoft Graph — Data Path Mapping** | **Admin — Graph Permission Approval** |

**Several web parts composed on one page:**

![Finished IT Operations dashboard with six SharePoint Smart Charts web parts](screenshots/showcase-it-ops-page.png)

---

## Installation (No Build Required)

The pre-built package is included in the repository. To deploy without installing Node.js or building anything:

1. Download `sharepoint/solution/sharepoint-smart-charts.sppkg` from this repository.
2. Upload it to your **SharePoint App Catalog** (SharePoint Admin Center → Advanced → App Catalog → Apps for SharePoint).
3. Check **Make this solution available to all sites** and click **Deploy**.
4. (Only if you plan to use the **Microsoft Graph** data source) approve the pending Graph permission request — see [Graph API Permissions](#graph-api-permissions).
5. On any modern SharePoint page, click **Edit** → **+** → search for **SharePoint Smart Charts**.

---

## Prerequisites (for Development Only)

| Requirement | Detail |
|---|---|
| **Node.js** | 18.x LTS (`>=16.13.0 <19.0.0`) |
| **npm** | 8+ |
| **Gulp CLI** | `npm install -g gulp-cli` |
| **SharePoint** | Online (Microsoft 365) |
| **SPFx** | 1.20 |
| **Permissions to deploy** | Site Owner or above |

---

## Development Setup

```bash
# Install dependencies (already done if you received this as a built project)
npm install --legacy-peer-deps

# Update serve.json with your SharePoint site URL
# Edit config/serve.json → change initialPage to your workbench URL:
# https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/workbench.aspx

# Start local dev server
gulp serve
```

Open the Workbench URL shown in the terminal. Add the **SharePoint Smart Charts** web part to the page.

---

## Build & Deploy

```bash
# 1. Build a production bundle
gulp bundle --ship

# 2. Package into a .sppkg file
gulp package-solution --ship

# Output: sharepoint/solution/sharepoint-smart-charts.sppkg
```

**Deploy to SharePoint:**
1. Go to **SharePoint Admin Center** → **App Catalog**
2. Upload `sharepoint-smart-charts.sppkg`
3. Check **Make this solution available to all sites** (or deploy to specific sites)
4. On any SharePoint page, click **Edit** → **+** → search for **SharePoint Smart Charts**

---

## Configuration

All settings are managed through the web part property pane (click the pencil icon while the page is in Edit mode). Data source selection, column mapping, data controls, and the Advanced Options panel live **inline in the web part** while editing, because they depend on the loaded data.

The pane is organized into three pages — **Chart**, **Appearance**, and **Advanced** — navigated with the Back / Next links at the bottom. Only settings relevant to the currently selected chart type are shown; options that don't apply are hidden automatically.

### Page 1 — Chart

**Header**

| Setting | Default | Description |
|---|---|---|
| **Show Header** | On | Display a prominent title above the chart |
| **Header Text** | _(blank)_ | The title text shown in the header |

**Chart Settings**

| Setting | Default | Description |
|---|---|---|
| **Chart Type** | Bar (Vertical) | 17 types organized in Standard and Specialized groups |
| **Chart Title** | _(blank)_ | Title rendered inside the chart canvas |
| **Legend Position** | Bottom | Top, Bottom, Left, or Right |
| **Chart Height** | 400 px | Height of the chart in pixels |
| **Show Legend** | On | Toggle the chart legend on/off |
| **Stacked** | Off | Stack multiple series (Bar and Line charts) |
| **Show Data Table** | Off | Display a paginated data table below the chart |
| **Show Export Bar** | On | Show PNG / JPEG / CSV / Excel export buttons |
| **X / Y Axis Label** | _(blank)_ | Labels along each axis |
| **Histogram Bins** | 10 | Number of bins (4–50) — Histogram chart type only |

### Page 2 — Appearance

| Setting | Default | Description |
|---|---|---|
| **Color Palette** | Office | 7 palettes: Office, Vibrant, Pastel, Monochrome, Traffic Light, Warm, Cool |
| **Show Data Labels** | Off | Annotate each data point with its value |
| **Value Prefix / Suffix** | _(blank)_ | Text around each label (e.g. `$`, `%`) |
| **Decimal Places** | 0 | Decimal places shown (0–4) |
| **Abbreviate Numbers (K/M)** | Off | Abbreviates 1,000 → 1K, 1,000,000 → 1M, 1,000,000,000 → 1B |
| **Y Axis Minimum / Maximum** | _(auto)_ | Override the Y axis range (validated numeric) |
| **Logarithmic Scale** | Off | Switch the Y axis to log scale |
| **Logarithmic Scale (X)** | Off | Switch the X axis to log scale — Scatter and Bubble charts only |
| **Step Interpolation** | Off | Render as a stepped line instead of straight/smooth segments — Line and Area charts only |
| **Show Grid Lines** | On | Toggle grid lines |
| **X Label Rotation** | 0° | Degrees to rotate X axis labels |
| **X Axis Type** | Auto-detect | Auto-detect / Category / Time (dates) |

### Page 3 — Advanced

| Setting | Default | Description |
|---|---|---|
| **Trendline** | None | None, Linear, or Moving Average overlay per series |
| **Moving Average Window** | 3 | Window size (2–20) for the moving average |
| **Forecast Periods** | 0 | Project the linear trendline up to 12 periods past the data |
| **Reference Line** | None | Fixed value, Mean, or Median dashed line |
| **Fixed Value / Line Color** | _(blank)_ / `#666666` | Value and color for the reference line |
| **Right-axis series** | _(blank)_ | Comma-separated column names plotted against a second, right-hand Y axis — Bar, Line, Area only |
| **Right axis label / Log Scale (Right Axis)** | _(blank)_ / Off | Caption and optional log scale for the right axis |
| **Error Bar Type** | None | None, Custom column (±), Std Dev (computed), or Std Error of Mean (computed) — Bar, Line, Area only |
| **Error value column** | _(blank)_ | Column holding the per-row ± margin — only used when Error Bar Type is Custom |
| **Bracket pairs** (Significance Brackets) | _(blank)_ | Newline-delimited `col1,col2,label` (or a JSON array) — draws an annotation bracket between two bar categories, e.g. `GroupA,GroupB,*` — Bar chart only |
| **Show Filters to Viewers** | Off | Filter bar for page visitors in view mode |
| **Details on Demand** | Off | Clicking a chart element shows its underlying rows |
| **Overlay Data Points on Bars** | Off | Show individual value markers on top of bars — Bar, Horizontal Bar only |
| **Show Bubble Size Legend** | Off | Small/medium/large size key for the Bubble chart |
| **Threshold Value / Direction / Color** | _(off)_ / Below / `#d13438` | Highlight values crossing a threshold |
| **Auto-Refresh Interval** | 0 (off) | Reload network sources every N minutes — disabled for the Upload File source |
| **Cache API Results** | 0 (off) | Cache REST/Graph responses for N minutes per session — only applies to the REST API and Microsoft Graph sources |

### Inline Advanced Options panel (while editing)

| Setting | Description |
|---|---|
| **Color by Column** | Color scatter/bubble points by a category column |
| **Tooltip Columns** | Extra columns appended to the hover tooltip |
| **Drill Levels 1–3** | Column hierarchy for click-to-drill navigation |
| **Bookmarks** | Save, apply, and delete named view states |

---

## Graph API Permissions

The package does **not** declare any Microsoft Graph permissions, so there is **no admin trust dialog** during deployment. The other four data sources (Upload File, SharePoint List, SharePoint File, REST API) and every chart feature work with no Graph permissions at all.

To use the **Microsoft Graph** data source, a tenant admin must grant the required scope(s) manually:

**Option A — Manual grant (no redeployment needed):**
1. Open **SharePoint Admin Center** → **Advanced** → **API access**.
2. Click **Manage** (or **+ Add a request**) and add a permission for the `SharePoint Online Client Extensibility Web Application Principal`.
3. Grant `User.Read` as a baseline. Add any additional scopes your Graph endpoints need (e.g. `Reports.Read.All`, `Group.Read.All`).

**Option B — Restore the package declaration (restores the standard trust-dialog flow):**
Add the following to `config/package-solution.json` under `"solution"` and redeploy:
```json
"webApiPermissionRequests": [
  { "resource": "Microsoft Graph", "scope": "User.Read" }
]
```
SharePoint will then prompt for admin approval during the next deployment, after which the permission is pre-approved for all Graph calls.

---

## Sample Data

The `sample-data/` folder contains ready-to-use files for testing the chart types (several are also provided as `.xlsx`):

| File | Best Chart Type | Columns |
|---|---|---|
| `monthly-sales.csv` | Bar, Line, Area, Waterfall | Month, Revenue, Units, Target, Profit |
| `regional-sales.csv` | Horizontal Bar, Treemap | Region, Revenue, Budget, Headcount |
| `market-share.csv` | Pie, Doughnut, Treemap | Product, MarketShare, Revenue, YoYGrowth |
| `scatter-rnd.csv` | Scatter | Company, RnDSpend, Revenue |
| `bubble-companies.csv` | Bubble | Company, Revenue, Employees, GrowthRate |
| `radar-products.csv` | Radar | Dimension, ProductA, ProductB, ProductC |
| `sales-by-month.csv` | Bar, Line, Area | Month, Online Sales, Store Sales, Total |
| `population-gdp.csv` | Scatter, Bubble, Histogram | Country, GDP Per Capita, Life Expectancy, Population (M) |
| `product-ratings.csv` | Radar | Product, Performance, Reliability, Ease of Use, Value, Support, Features |
| `study-hours-scores.csv` | Scatter, Histogram, Box Plot | Study Hours, Test Score, Confidence Level |
| `temperature-trend.csv` | Line, Area | Month, High Temp (F), Low Temp (F), Avg Temp (F), Rainfall (in) |
| `api-response-example.json` | Bar, Line | Quarter, Revenue, Expenses, Profit |

To run a local REST API test server:
```bash
node sample-data/test-api-server.js
# API available at http://localhost:3001/data  (Data Path: value)
```

---

## Project Structure

```
sharepoint-smart-charts/
├── config/
│   ├── config.json               # Bundle entry points + localization
│   ├── package-solution.json     # Solution ID, version
│   └── serve.json                # Local dev server config
├── mockups/                      # Screenshot guide HTML
├── sample-data/                  # Test data files (not deployed)
├── screenshots/                  # Documentation screenshots
├── src/
│   └── webparts/
│       └── sharePointSmartCharts/
│           ├── SharePointSmartChartsWebPart.ts      # Web part class, property pane, Dynamic Data source
│           ├── SharePointSmartChartsWebPart.manifest.json
│           ├── components/
│           │   ├── SharePointSmartCharts.tsx        # Root component — pipeline, drill-down, bookmarks
│           │   ├── DataSourcePanel.tsx               # Data loading UI (5 sources, sheet picker)
│           │   ├── ColumnMapper.tsx                  # Column → axis mapping, per-series colors/types
│           │   ├── DataControls.tsx                  # Sort, filter, limit, group-by aggregation
│           │   ├── AdvancedOptions.tsx               # Color-by, tooltips, drill hierarchy, bookmarks
│           │   ├── ChartRenderer.tsx                 # Chart.js rendering — all 17 chart types
│           │   ├── DataTable.tsx                     # Tabular data view
│           │   ├── ExportBar.tsx                     # PNG / JPEG / CSV / Excel export
│           │   └── SharePointSmartCharts.module.scss
│           ├── services/
│           │   └── dataLoaders.ts                    # Shared loaders: list, file, REST, Graph, cache
│           ├── types/index.ts                        # Shared TypeScript types
│           └── loc/                                  # Localization strings
├── CHANGELOG.md
├── USER-GUIDE.md
├── package.json
└── tsconfig.json
```

---

## Key Dependencies

| Package | Purpose |
|---|---|
| `@microsoft/sp-webpart-base` | SPFx 1.20 web part base |
| `@microsoft/sp-dynamic-data` | Publishes chart-click selections to connected web parts |
| `chart.js` + `react-chartjs-2` | Core chart rendering |
| `chartjs-plugin-datalabels` | On-chart value labels |
| `chartjs-adapter-date-fns` + `date-fns` | Time-scale X axis |
| `@sgratzl/chartjs-chart-boxplot` | Box Plot chart type |
| `chartjs-chart-treemap` | Treemap chart type |
| `chartjs-chart-matrix` | Heatmap chart type |
| `papaparse` | CSV parsing |
| `xlsx` 0.20.3 | Excel parsing (official SheetJS build, Apache 2.0 — patched for CVE-2023-30533 / CVE-2024-22363) |
| `@pnp/sp` | SharePoint list/file access via PnPjs v3 |

---

## Troubleshooting

- **"Scatter and bubble charts need numeric values…"** — the mapped X or Y column contains no numbers. Pick a numeric column in the Column Mapping panel.
- **Microsoft Graph source returns 401/403** — the Graph permission request has not been approved, or your endpoint needs a scope not listed in `package-solution.json`. See [Graph API Permissions](#graph-api-permissions).
- **"This list has 5,000 or more items…"** — SharePoint returns at most 5,000 items per request. Filter the list with a view-backed approach or aggregate the data upstream.
- **Arrows or accented characters look garbled in an exported CSV** — fixed in 1.1.0; exports now include a UTF-8 BOM. Re-export with the current version.
---

## Limitations

- **Upload file data size:** Uploaded file data is serialized to the web part property bag. Datasets up to 200 KB persist across page reloads. Larger files display for the current session only — store them in a SharePoint document library and use the SharePoint File source for fully persistent large datasets.
- **SharePoint list size:** A maximum of 5,000 items is loaded per list (SharePoint REST limit); a warning is shown when this cap is hit.
- **Microsoft Graph:** Requires tenant-admin approval of Graph permissions; without approval the Graph source cannot load data.
- **Direct database connections:** SPFx cannot connect directly to databases. Use a REST API (e.g., Azure Function or custom API) that queries your database and returns JSON.
- **Cross-origin REST APIs:** The browser's same-origin policy applies. For external APIs, ensure CORS headers are configured on the server.
- **Forecast:** Trendline forecasting extends category axes only (not the time axis).

---

## License

[MIT](LICENSE) © 2026 Sean Regan
