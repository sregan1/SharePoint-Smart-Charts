# Code Review Fix Plan — July 2026

Execution plan from the July 10, 2026 code review of SharePoint Smart Charts v1.2.0.
Work through items in order (Phase 1 → 4), checking boxes as you complete them.

## Context for the executor

- **Project**: SPFx 1.20 web part, React 17 (functional components + hooks), TypeScript 4.7, Chart.js 4.x + react-chartjs-2 5.x. Plugins: chartjs-plugin-datalabels, chartjs-chart-treemap, chartjs-chart-matrix, @sgratzl/chartjs-chart-boxplot, chartjs-adapter-date-fns.
- **Build/verify**: `npm run build` (gulp bundle — runs tsc, ESLint, webpack). Baseline is clean as of this review. Run it after each phase at minimum; a phase is not done until it builds clean.
- **Line numbers** below were accurate at review time (commit `f208b01`). Verify context before editing — earlier fixes in this plan will shift later line numbers.
- **Critical codebase invariant**: `onPropertiesUpdate` mutates `this.properties` on the web part WITHOUT re-rendering React. Therefore (a) any inline-edited setting must flow through component STATE (props go stale until the next web-part render), and (b) every inline handler that changes persisted config must update BOTH state and properties. Several bugs below are violations of this invariant — do not introduce new ones.
- **Localization**: all UI strings live in `src/webparts/sharePointSmartCharts/loc/en-us.js` + `mystrings.d.ts` (currently 257 keys, in parity — keep it that way; add both sides for any new string). Use American English.
- **Known tooling gotcha**: the Edit tool normalizes the literal 6-char escape `﻿` into an invisible BOM character. If you touch the BOM line in `handleExportCsv` (ChartRenderer.tsx ~line 319), verify the source still contains the escape sequence, not a raw BOM.

Key files (all under `src/webparts/sharePointSmartCharts/`):

| File | Role |
|---|---|
| `components/ChartRenderer.tsx` (~1400 lines) | All Chart.js data building, options, plugins, export |
| `components/SharePointSmartCharts.tsx` (~850 lines) | State, data pipeline (filter → aggregate → sort → limit), auto-load/refresh, drill-down, bookmarks |
| `SharePointSmartChartsWebPart.ts` (~640 lines) | Property pane (simple/advanced), dynamic data |
| `components/DataSourcePanel.tsx` | Data source UI + manual loads |
| `services/dataLoaders.ts` | All loaders + sessionStorage cache |
| `components/ColumnMapper.tsx` | Column mapping + per-series color/type |
| `components/AdvancedOptions.tsx` | Color-by, tooltips, drill levels, bookmarks UI |
| `types/index.ts` | Shared types, palettes, extractColumns, parseBookmarks |

---

## Phase 1 — High-severity bugs

### [x] 1.1 Chart-type effect silently rewrites a valid numeric X mapping (data loss)

`components/SharePointSmartCharts.tsx:293-321` — the `React.useEffect(..., [chartType])` "switching FROM a numeric-X chart" `else` branch (lines 309-318) runs on **first mount** too (for upload sources, the lazy initializer at lines 167-223 restores `uploadedData` synchronously, so `state.data.length > 0` on mount). It cannot distinguish "user just left scatter" from "user deliberately mapped a numeric column."

**Failure**: bar chart of uploaded CSV with `xColumn = "Year"` (numeric strings count — `isNumericCol` matches any parseable value). Every page load in edit mode remaps X to the first non-numeric column and **persists it** via `onPropertiesUpdate`. Also fires on bar → line switches with numeric X.

**Fix**: track the previous chart type in a ref. Run the `else` (reset) branch only when the previous type was in `NUMERIC_X_TYPES` and the new one is not; skip the effect body entirely on first mount:

```tsx
const prevChartTypeRef = React.useRef<string | undefined>(undefined);
React.useEffect(() => {
  const prevType = prevChartTypeRef.current;
  prevChartTypeRef.current = chartType;
  if (prevType === undefined || prevType === chartType) return; // first mount / no change
  ...
  } else {
    if (NUMERIC_X_TYPES.indexOf(prevType) < 0) return; // wasn't a numeric-X chart before
    ...existing reset logic...
  }
}, [chartType]);
```

**Verify**: with data present and a numeric X on a bar chart, remount (reload workbench) → mapping unchanged; scatter → bar still resets X; bar → line leaves numeric X alone.

### [x] 1.2 `count` aggregation always renders a blank chart

`components/SharePointSmartCharts.tsx:75-77` (producer) + `effectiveColumnConfig` (~line 598). `aggregateRows` with `count` emits rows containing only `[groupByColumn, Count]`, but `columnConfig.yColumns` still points at original columns → ChartRenderer reads all-null → blank chart, no error. The `Count` column is **unselectable**: ColumnMapper options come from `state.columns` (pre-aggregation), and `buildColumnConfig`/`handleDataLoaded` strip any column not in the original list.

**Fix**: in `effectiveColumnConfig`, when the effective aggregation is `count`, override `yColumns: ['Count']` (and X to the effective group-by column when set). Mirror this in the drill-down path if it aggregates with count.

**Verify**: load data, set Group by + Aggregation = Count in DataControls → chart shows counts per group. Also test drill-down with count.

### [x] 1.3 JPEG export is black; PNG export invisible in dark theme

`components/ChartRenderer.tsx:300-308` — canvas is never given an opaque background. JPEG drops alpha (transparent → black); dark-theme PNG is near-white text on transparency (invisible on white viewers).

**Fix**: composite onto an offscreen canvas before export (preferred — doesn't affect on-screen rendering):

```tsx
const exportImage = (mime: string, quality: number, filename: string) => {
  const src = chartRef.current?.canvas as HTMLCanvasElement | undefined;
  if (!src) return;
  const out = document.createElement('canvas');
  out.width = src.width; out.height = src.height;
  const ctx = out.getContext('2d');
  if (!ctx) return;
  ctx.fillStyle = isDarkTheme ? '#1b1a19' : '#ffffff';
  ctx.fillRect(0, 0, out.width, out.height);
  ctx.drawImage(src, 0, 0);
  downloadUrl(out.toDataURL(mime, quality), filename);
};
```

**Verify**: export JPEG in light theme (white background, readable) and PNG in dark theme (dark background, readable when pasted on white).

### [x] 1.4 Click handler / tooltip extras map element index to the wrong source row

`components/ChartRenderer.tsx:463-476` (`handleChartClick` uses `data[index]`) and `:482-490` (`afterBody` uses `data[items[0].dataIndex]`). Several builders filter rows, so element index ≠ source row index:

- **Pie/doughnut** (~line 811): rows with null Y are filtered out of `validRows` → clicking a slice emits the wrong category to Dynamic Data consumers; tooltip extras show the wrong row.
- **Scatter/bubble** (~lines 749-778): non-numeric points are dropped, shifting indexes. The `colorByColumn` branch is doubly wrong: datasets are per-category subsets, so `index` indexes the subset, and the `datasetIndex >= validYColumns.length` guard (line 466) blocks clicks on all but the first category.
- **Time axis** (~lines 581-600): `toPoints` drops unparseable-date rows; both click indexes and the per-point threshold color arrays (built over ALL rows) misalign with the filtered points — threshold coloring highlights wrong points too.

**Fix**: carry the source row index through every filtered builder — e.g. build points as `{ x, y, _rowIndex }` (Chart.js parsing tolerates extra keys on point objects; for pie keep a parallel `validRowIndexes: number[]`). In `handleChartClick` and `afterBody`, resolve via the carried index when present, falling back to the raw element index for unfiltered builders. For the time-axis threshold arrays, build colors from the same filtered point list, not from `data`. For `colorByColumn`, resolve per-dataset (each dataset keeps its own row-index array) and fix the `datasetIndex` guard so category datasets are clickable.

**Verify**: pie with a null-Y row → clicking each slice emits its own category; scatter with some non-numeric rows + tooltipColumns → hover shows the hovered row's values; time axis with threshold coloring + one bad date → correct points highlighted.

### [x] 1.5 Graph API results silently truncated at one page

`services/dataLoaders.ts:103-111` — `loadGraphApi` issues a single `.get()` and ignores `@odata.nextLink`; most Graph collections page at 100 items. Returns `truncated: false`, so charts are quietly wrong with no warning.

**Fix**: follow `@odata.nextLink` in a loop (cap at ~10 pages / 5,000 rows to match `SP_LIST_ROW_LIMIT`), concatenating each page's `value`. Unwrap `value` per page BEFORE `applyDataPath` semantics get involved (apply `dataPath` to the first page's envelope only if it's not the standard `value` path — simplest correct approach: if the response has `@odata.nextLink`/`value`, paginate on `value` and apply a non-`value` `dataPath` per item set as today). Set `truncated: true` when stopping at the cap with a `nextLink` remaining, so the existing truncation warning UI fires.

**Verify**: build passes; hand-test logic against a mocked two-page response shape (write a tiny node script if useful). Confirm `truncated` reaches the DataSourcePanel warning for the graphApi source type.

---

## Phase 2 — Medium-severity bugs

### [x] 2.1 Auto-refresh / manual refresh read stale props after inline edits

`components/SharePointSmartCharts.tsx:234-288` (auto-load effect), `:324-330` (interval), and `handleRefresh` (~line 532). All read `props.dataSourceType/dataUrl/listName/...`, which go stale after inline edits (invariant above). Failure: after changing the REST URL inline, an auto-refresh tick refetches the OLD URL and overwrites fresh data; `handleRefresh` clears the wrong cache key.

**Fix**: read source config from `state.dataSourceConfig` (already mirrored and fresh) in all three places. Cleanest: keep a `dataSourceConfigRef` in sync (like `columnConfigRef`) and read it inside the effect/handlers, keeping `refreshKey` as the only trigger dep. The cache key in both the effect and `handleRefresh` must be built from the same fresh values.

### [x] 2.2 Applying a bookmark in edit mode never persists

`components/SharePointSmartCharts.tsx:488-513` — `handleApplyBookmark` updates state mirrors + `columnConfigRef` but never calls `onPropertiesUpdate`. Author applies a bookmark, saves the page → published viewers get the pre-bookmark chart; a later single-field save persists a mixed state that never existed on screen.

**Fix**: when `!isReadOnly`, also call `onPropertiesUpdate` with all applied fields (sortColumn, sortDirection, filterColumn, filterValue, groupByColumn, aggregation, xColumn, yColumns). Keep read-mode apply non-persisting (it has no `onPropertiesUpdate` effect anyway, but make the intent explicit).

### [x] 2.3 Numeric strings sort lexicographically

`components/SharePointSmartCharts.tsx:573-580` — comparator only compares numerically when both values are `typeof number`; string-encoded numbers (REST APIs, SP text columns) sort "100" < "20" < "9".

**Fix**: coerce first, mirroring the aggregation coercion at line 87:

```ts
const an = Number(av), bn = Number(bv);
if (av !== '' && bv !== '' && !isNaN(an) && !isNaN(bn)) return dir * (an - bn);
```

### [x] 2.4 Excel date columns load as raw serial numbers

`services/dataLoaders.ts:34-42` — `XLSX.read` without `cellDates` + default `raw: true` returns dates as serials (45329 instead of 2024-02-07), so date columns chart as meaningless numbers and the time axis can't work for Excel sources.

**Fix**: `XLSX.read(buffer, { type: 'array', cellDates: true })`, then in `sheetToRows` convert `Date` instances to ISO date strings (`d.toISOString().slice(0, 10)`, or include time when non-midnight). Do NOT leave raw `Date` objects in rows — `extractColumns` (types/index.ts ~line 96) excludes object-valued columns, so date columns would vanish entirely. Apply to both `parseExcelBuffer` (upload) and `loadSharePointFile` (URL) paths.

**Verify**: parse a real .xlsx with a date column (script it with the project's own xlsx dep) → rows contain ISO strings; column appears in mapper; time axis works.

### [x] 2.5 Per-series color/type overrides desync when a middle Y column is unchecked

`components/ColumnMapper.tsx:54-82` — overrides are positional comma-joined lists, but `toggleYColumn` (76-82) removes a column without splicing its override slot: uncheck the middle of [Sales, Cost, Profit] → Profit inherits Cost's color/type. Also `getSeriesColor` (54-57) falls back to hardcoded `#0078d4` while the chart actually uses palette rotation — swatches lie for series 2+.

**Fix**: in `toggleYColumn`, when removing, splice the same index out of both override arrays and emit via `onSeriesColorsChange`/`onSeriesTypesChange`; when adding, no change needed (append semantics already align). Default the swatch to the palette color: `PALETTES[palette][yIdx % PALETTES[palette].length]` (pass the palette name down — check what props ColumnMapper already receives). 

### [x] 2.6 Delimiter dropdown does nothing for an already-uploaded text file

`components/DataSourcePanel.tsx` (~lines 139, 252-254, 313-327, 448) — raw CSV text isn't retained after parse (only `workbookRef` for Excel), and the Load button is hidden once loaded, yet the delimiter select shows exactly then. Changing it has no effect.

**Fix**: keep the raw text in a `rawTextRef` (mirror of `workbookRef`); on delimiter change while an uploaded text file is active, re-run `parseCsvText(rawTextRef.current, newDelimiter)` and `handleDataLoaded`. Persist the delimiter as today.

### [x] 2.7 Bad `dataPath` (or null rows) crashes with a cryptic TypeError

`services/dataLoaders.ts:80-89` + `types/index.ts` `extractColumns` (~line 91). A typo'd path resolves to `undefined` → `[undefined]` → `Object.keys(undefined)` throws "Cannot convert undefined or null to object" into the error box.

**Fix**: in `applyDataPath`, when the resolved value is `undefined`/`null`, throw `new Error(...)` with a localized message naming the path (add a loc string, e.g. `DataPathNotFoundError` with `{0}` placeholder via the existing `fmt()` helper). In `extractColumns`, skip entries where `rows[i]` is not a non-null object.

### [x] 2.8 SharePoint URL returning HTML parses as CSV "successfully"

`services/dataLoaders.ts:67-77` — sharing links (`/:x:/r/...`), login redirects, etc. return `200 text/html`; it parses into garbage rows and reports success.

**Fix**: in the text branch, if `response.headers.get('content-type')` contains `text/html` OR the trimmed text starts with `<!doctype`/`<html` (case-insensitive), throw a localized error telling the user to use a direct file path, not a sharing link. Add the loc string.

### [x] 2.9 Duplicate CSV headers silently drop a column — VERIFY FIRST

`services/dataLoaders.ts:21-32` — claim: PapaParse with `header: true` doesn't dedupe repeated headers, so the later column overwrites the earlier per row. **First write a 5-line node script against the installed papaparse to confirm the behavior for this version** (recent PapaParse versions may auto-rename duplicates — e.g. `Amount_1`). If confirmed: pass a `transformHeader: (h, i)` that suffixes duplicates (`Amount`, `Amount_1`), matching SheetJS behavior. If PapaParse already dedupes, check the box with a note and move on.

**Verified 2026-07-14**: installed papaparse is 5.5.3, which already dedupes (`Amount,Region,Amount` → keys `Amount`, `Region`, `Amount_1`, logs "Duplicate headers found and renamed."). No fix needed.

### [x] 2.10 `beforeAfter` and `violin` charts emit garbage Dynamic Data selections

`components/ChartRenderer.tsx:461-462` — `clickableType` excludes histogram/boxplot/treemap/heatmap but not `violin` (per-group elements → `data[k]` is an arbitrary row) or `beforeAfter` (datasets are per-row, index is 0/1 → wrong row AND wrong series). Tooltip extras are equally wrong for these types (see 3.3).

**Fix**: add `'violin'` and `'beforeAfter'` to the non-clickable list.

### [x] 2.11 Dual axis on horizontalBar is broken

`components/ChartRenderer.tsx:495-496` includes `horizontalBar` in `hasDualAxis` and ~line 616 assigns `yAxisID: 'y1'`, but the horizontal options define no `y1` scale — and with `indexAxis: 'y'` the value axis is X anyway, so Chart.js auto-creates a category `y1` and the series renders wrong/blank. `supportsDualAxis` in `types/index.ts` (~line 198) advertises it, so it's reachable from the pane.

**Fix (simplest correct)**: remove `horizontalBar` from both `hasDualAxis` (ChartRenderer) and `supportsDualAxis` (types/index.ts), so the pane hides the option for it. (A proper second-x-scale implementation is optional; not required.)

### [x] 2.12 Log-scale X is wired wrong for scatter/bubble and breaks histogram

`components/ChartRenderer.tsx` — two defects: (a) ~line 1119 `scatterOptions` sets the **x** axis type from `logScale` (the Y toggle), so "Log Y" logs both axes on scatter/bubble while the actual `logScaleX` prop does nothing there; (b) `logScaleXApplies` (~lines 508-512) includes `'histogram'`, whose X values are category bin labels → `type: 'logarithmic'` parses them as NaN → blank chart.

**Fix**: in `scatterOptions`, x axis type from `logScaleX`, y axis type from `logScale`. Remove `'histogram'` from `logScaleXApplies`. Check the property pane (`SharePointSmartChartsWebPart.ts`) so the Log X toggle isn't offered for histogram if it's conditionally shown.

### [x] 2.13 3-digit hex palette entries produce invalid colors when alpha is concatenated

`types/index.ts:76` — `monochrome` palette ends with `'#555'`, `'#333'`; builders concat alpha suffixes (`` `${color}cc` `` → `'#555cc'`, invalid, canvas silently ignores). Reachable with monochrome + ≥9 series/slices, or any user `seriesColors` entry that isn't 6-digit hex.

**Fix**: (a) change the two palette entries to `'#555555'`/`'#333333'`... wait — those duplicate existing entries; use distinct values like `'#444444'` and `'#666666'` (keep 10 distinct grays). (b) Harden centrally: in `resolveColors` (types/index.ts ~line 142), normalize each color — expand `#abc` → `#aabbcc`; leave anything else as-is. (c) Where alpha is appended in ChartRenderer, only append when the color matches `^#[0-9a-fA-F]{6}$`, else use the color as-is (a shared `hexWithAlpha(color, alphaHex)` helper — one may already exist; search for `hexWithAlpha` and reuse/extend it).

### [x] 2.14 Linear trendline on a time axis regresses over row index, not time

`components/ChartRenderer.tsx:639-652` + `linearTrend` (~line 186). On `xIsTime`, the least-squares fit is over 0..n-1 but plotted at real timestamps → kinked/wrong trend for irregular intervals (typical SP list data); rows with unparseable dates are included in the fit but dropped from the plot.

**Fix**: when `xIsTime`, regress y against parsed timestamps over only the rows with valid timestamps, then plot `m*t + b` at those timestamps. Keep the index-based fit for category axes. If a forecast is active on a time axis, extend by the median timestamp interval per period.

---

## Phase 3 — Low-severity bugs

### [x] 3.1 Pie/doughnut legend loses dark-theme text color

`components/ChartRenderer.tsx:~1029` — `pieOptions` replaces the whole `legend` object, omitting `labels: { color: textColor }`. Add it back.

### [x] 3.2 Significance brackets silently never render on horizontalBar

`components/ChartRenderer.tsx:~1221-1262` — plugin calls `xScale.getPixelForValue(categoryLabel)` on a linear scale (horizontal layout swaps axes) → NaN, draws nothing. **Fix**: restrict the feature to vertical `bar` (both the plugin gate and the property pane visibility if conditional).

### [x] 3.3 Tooltip extras show nonsense rows on binned/derived chart types

`components/ChartRenderer.tsx:482-490` — `afterBody` indexes `data[]` by bin/group index on histogram/boxplot (and violin/beforeAfter). **Fix**: skip the `afterBody` callback for `histogram`, `boxplot`, `violin`, `beforeAfter` (reuse or parallel the non-clickable list from 2.10 — consider one shared `rowMappedType` constant).

### [x] 3.4 Error-bar length wrong on logarithmic value axes

`components/ChartRenderer.tsx:~1179-1181` — pixel length computed as `|px(err) − px(0)|`, meaningless on log scales. **Fix**: compute per point `|px(v + err) − px(v)|` for the up-whisker and `|px(v) − px(max(v − err, tiny))|` for the down-whisker; or simply skip drawing error bars when the value axis is logarithmic (acceptable if simpler).

### [x] 3.5 `cacheMinutes` offered for all source types but only works for REST/Graph

`SharePointSmartChartsWebPart.ts:~628-634` — disable the slider unless `dataSourceType` is `restApi`/`graphApi` (same pattern as `refreshIntervalMinutes` for upload), and/or note it in the field description. Do not implement SP list/file caching.

### [x] 3.6 Background auto-refresh slams the config panel shut mid-edit

`components/SharePointSmartCharts.tsx:~363` — `handleDataLoaded` forces `isConfigOpen: false`; when the auto-refresh interval triggers a reload while the author has the panel open, it closes under them. **Fix**: add an optional flag (e.g. `handleDataLoaded(data, columns, { fromAutoLoad?: boolean })`) and preserve `prev.isConfigOpen` for auto/refresh loads.

### [x] 3.7 `parseBookmarks` accepts malformed entries that crash on apply

`types/index.ts:~125-133` — `["a"]` or `[{"name":"x"}]` parse fine, then `bookmark.state.sortColumn` throws in the read-mode dropdown, unmounting the web part. **Fix**: filter parsed entries to those with a string `name` and a non-null object `state`.

### [x] 3.8 `Count` column name is hardcoded English and can collide with a real column

`components/SharePointSmartCharts.tsx:~76`. Two parts: (a) if the group-by column is itself named `Count`, `out.Count = members.length` destroys the group labels — disambiguate (e.g. use `Count of rows` or append a suffix when colliding). (b) Localization of the name is OPTIONAL and touchy — the string becomes a column key that 1.2 (item 1.2 above) must match; if you localize it, define ONE shared constant used by both `aggregateRows` and `effectiveColumnConfig`. Simplest safe fix: keep `Count` as the key, fix only the collision, and export a `COUNT_COLUMN` constant both sites import.

### [x] 3.9 Manual panel loads bypass the cache — stale data resurrects on reload

`components/DataSourcePanel.tsx:~191-217` — after a successful manual Load Data for restApi/graphApi, the sessionStorage cache still holds the older entry; next page load serves it. **Fix**: after a successful manual load, clear (simplest) the cache entry for the same key the auto-load effect uses. Keep key construction in ONE exported helper in `dataLoaders.ts` (e.g. `buildCacheKey(srcType, dataUrl, dataPath)`) so the panel, effect, and `handleRefresh` (2.1) can't drift.

### [x] 3.10 List discovery fires a PnPjs request per keystroke in Site URL

`components/DataSourcePanel.tsx:~76-100` — debounce ~500 ms (setTimeout in the effect + clear in cleanup; the existing `cancelled` flag already handles stale results).

### [x] 3.11 Sheet selector persists after switching data-source type

`components/DataSourcePanel.tsx:~264-265, 329-343` — clear `sheetNames` state and `workbookRef` (and `rawTextRef` from 2.6) in the source-type button handler.

### [x] 3.12 Clearing drill Level 1 leaves orphaned lower levels

`components/AdvancedOptions.tsx:~55-62` — `setDrillLevel(0, '')` yields `",B,C"`. **Fix**: when a level is cleared, truncate the array at that level (`next.length = level`) before joining.

### [x] 3.13 `.xlsm`/`.xlsb` URLs parsed as CSV garbage

`services/dataLoaders.ts:44-47` — add `.xlsm` and `.xlsb` to `isExcelUrl`, and to the upload accept/extension check in `DataSourcePanel.tsx` (~lines 142, 280).

### [x] 3.14 Packaging metadata inconsistencies

`config/package-solution.json` — `features[0].version` is `1.1.0.0` while the solution is `1.2.0.0` (bump the feature version, and keep them in lockstep in future releases); `developer.websiteUrl` uses `http://` — change to `https://`.

---

## Phase 4 — Improvements (non-bugs; do after all bug phases build clean)

### [x] 4.1 Memoize ChartRenderer — PARTIAL, scope deliberately reduced

Biggest perf win. Every parent render (e.g. each keystroke in the viewer filter) rebuilds all datasets/options with new identities, forcing full `chart.update()`. Wrap the chart data builders and options objects in `React.useMemo` with correct deps, wrap the component in `React.memo`, and hoist a single `Intl.NumberFormat` instance out of `formatValue` (it's called per data label per draw). Be careful: deps lists must include every prop/state the builders read — an incorrect deps list here is worse than no memo. Also debounce the read-mode viewer filter input (~250 ms) in `SharePointSmartCharts.tsx` (~line 699).

**Done**: `ChartRenderer` wrapped in `React.memo` (no custom comparator — default shallow prop compare); `formatValue` now caches one `Intl.NumberFormat` per distinct `decimals` value instead of reconstructing per label per draw; the read-mode viewer filter value is debounced 250ms before it feeds `filteredRows`'s `useMemo` (the `<input>` itself stays instant — only the expensive filter/aggregate/chart-rebuild is delayed).

**Deliberately NOT done**: wrapping every one of ChartRenderer's ~15 dataset/option builders in individual `useMemo`s with hand-verified dependency lists. That rewrite touches nearly the whole file, and an incorrect deps array reintroduces stale-closure bugs (the same class of bug fixed in 1.1) with no automated test suite or running SharePoint workbench in this environment to visually verify chart output against. `React.memo` alone gives a partial win today (skips rebuilding on renders where `data`/`columnConfig` keep the same identity — true whenever count-aggregation/drill-down aren't active, since `processedData` is already `useMemo`'d in the parent) without that risk. Revisit the full builder-level memoization only with a way to visually verify each chart type's output (workbench or screenshot diffing) before/after.

### [x] 4.2 Replace O(n·k) category grouping with single-pass Maps

`components/ChartRenderer.tsx` (~lines 740, 747, 783, 790, 983-984) — `categories.indexOf` + per-category `data.filter` → build a `Map<string, rows[]>` in one pass (colorBy partitioning, heatmap, and category discovery).

### [x] 4.3 Radar ignores yAxisMin/yAxisMax

`components/ChartRenderer.tsx:~1139-1149` — pass `min`/`max` to the `r` scale like the other chart types.

### [x] 4.4 Detect Excel serial dates under forced time axis

If a user forces `xAxisType: 'time'` and X is all-numeric (Excel serials that predate fix 2.4, or any numeric column), `toPoints` drops every row → blank chart, no message. Show the existing mapping-warning path naming the column, or convert plausible serials (`(serial − 25569) * 86400000` for values in ~15000-60000).

### [x] 4.5 Export filenames from chart title

`components/ChartRenderer.tsx` — derive `chart.png`/`data.csv`/etc. from `chartTitle` (sanitize to safe filename chars, fall back to current names). Also add a `B` tier to `formatValue` abbreviation (currently 1.5e9 → "1,500M").

### [x] 4.6 Filter SharePoint internal fields from list loads

`services/dataLoaders.ts:57` uses `select('*')` — internal fields (`OData__UIVersionString`, `ContentTypeId`, `GUID`, `FileSystemObjectType`, …) flood the column dropdowns. Add a post-load exclusion list of known internal names/prefixes in `extractColumns` or the list loader.

### [x] 4.7 Surface partial CSV parse errors

`services/dataLoaders.ts:28-30` — when `result.errors.length > 0` but rows parsed, return an error count so DataSourcePanel can show a warning ("N rows could not be parsed") via the existing warning banner. Requires threading a `parseWarnings`-style field through `ILoadResult`.

### [x] 4.8 REST/file fetch timeout + JSON content-type guard

`services/dataLoaders.ts` — add `AbortController` with ~30 s timeout to `loadRestApi`/`loadSharePointFile`; on non-JSON content-type in `loadRestApi`, throw a targeted localized error instead of letting `response.json()` produce "Unexpected token <".

### [x] 4.9 Adopt the existing union types in the props interface

`components/ISharePointSmartChartsProps.ts` declares `sortDirection`, `aggregation`, `xAxisType`, `trendline`, `referenceLineType`, `legendPosition`, `thresholdDirection` as `string`; `types/index.ts` already defines matching union types. Adopt them and remove the `as any` cast at `SharePointSmartCharts.tsx:~529`. Expect some ripple through property-pane defaults — keep the change mechanical.

### [x] 4.10 Validate color fields in the property pane

`SharePointSmartChartsWebPart.ts:~506-509, 610-613` — `referenceLineColor`/`thresholdColor` accept any string. Add `onGetErrorMessage` accepting `#rgb`/`#rrggbb`/empty (reuse one validator). Also remove `stepLine` from `PANE_STRUCTURE_FIELDS` (~line 210) — it forces pane rebuilds nothing depends on.

---

## Final verification

1. [x] `npm run build` clean (tsc + lint + webpack). Verified 2026-07-14 after every phase and again at the end — clean throughout.
2. [x] Loc parity: `en-us.js` and `mystrings.d.ts` key sets identical — 265/265 keys (up from 257; 8 new strings added across 1.5, 2.7, 2.8, 4.4, 4.7, 4.8, 4.10, all added to both files).
3. [x] Grep for `data[index]`/`data[items[0].dataIndex]` in ChartRenderer — only the intentional fallback (`return data[index];`) inside `resolveSourceRow` remains, for chart types that don't filter rows (bar/line/area non-time, radar, waterfall). No stray raw-index lookups.
4. [ ] Sanity-run the workbench if available: bar chart with numeric X survives reload (1.1); count aggregation renders (1.2); exports readable in both themes (1.3). **NOT done** — no running SharePoint workbench in this environment; only `npm run build` (tsc/lint/webpack) was verified. Manually test before shipping.
5. [x] Did NOT bump the package version or commit. `git status` shows all changes as uncommitted working-tree modifications (plus this plan file, untracked).
