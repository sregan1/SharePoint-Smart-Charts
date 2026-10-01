import * as strings from 'SharePointSmartChartsWebPartStrings';

export type ChartType =
  | 'bar'
  | 'horizontalBar'
  | 'line'
  | 'area'
  | 'scatter'
  | 'pie'
  | 'doughnut'
  | 'bubble'
  | 'radar'
  | 'kpi'
  | 'histogram'
  | 'waterfall'
  | 'boxplot'
  | 'violin'
  | 'treemap'
  | 'heatmap'
  | 'beforeAfter';

export type ReferenceLineType = 'none' | 'fixed' | 'mean' | 'median';

export type DataSourceType =
  | 'upload'
  | 'sharePointList'
  | 'sharePointFile'
  | 'restApi'
  | 'graphApi';

export type AggregationType = 'none' | 'sum' | 'avg' | 'count' | 'min' | 'max';

export type XAxisType = 'auto' | 'category' | 'time';

export type TrendlineType = 'none' | 'linear' | 'movingAverage';

export type SortDirection = 'asc' | 'desc';

export type LegendPosition = 'top' | 'bottom' | 'left' | 'right';

export type ThresholdDirection = 'above' | 'below';

export interface IDataSourceConfig {
  dataSourceType: DataSourceType;
  uploadedFileName: string;
  siteUrl: string;
  listName: string;
  dataUrl: string;
  dataPath: string;
  delimiter: string;
  sheetName: string;
}

export interface IColumnConfig {
  xColumn: string;
  yColumns: string[];
  labelColumn: string;
  sizeColumn: string;
}

export interface IChartRecord {
  [key: string]: string | number | boolean | null | undefined;
}

export const CHART_COLORS: string[] = [
  '#0078d4',
  '#00b4d8',
  '#107c10',
  '#ffb900',
  '#d13438',
  '#8764b8',
  '#038387',
  '#e3008c',
  '#004578',
  '#69797e',
];

export const PALETTES: Record<string, string[]> = {
  // Okabe-Ito: distinguishable under the common forms of color blindness
  colorblind:  ['#0072b2','#e69f00','#009e73','#d55e00','#56b4e9','#cc79a7','#f0e442','#000000','#999999','#332288'],
  office:      ['#0078d4','#00b4d8','#107c10','#ffb900','#d13438','#8764b8','#038387','#e3008c','#004578','#69797e'],
  vibrant:     ['#e63946','#f4a261','#2a9d8f','#457b9d','#e9c46a','#264653','#a8dadc','#f77f00','#023e8a','#9b2226'],
  pastel:      ['#a8d8ea','#aa96da','#fcbad3','#ffffd2','#b5ead7','#ffdac1','#c7ceea','#e2f0cb','#ffb7b2','#ff9aa2'],
  monochrome:  ['#2d2d2d','#555555','#777777','#999999','#aaaaaa','#bbbbbb','#cccccc','#dddddd','#444444','#eeeeee'],
  trafficLight:['#107c10','#bad80a','#ffb900','#f7630c','#d13438','#647687','#008299','#0078d4','#69797e','#323130'],
  warm:        ['#d13438','#e74856','#f7630c','#ca5010','#ffb900','#f0a30a','#da3b01','#ef6950','#fce100','#fff100'],
  cool:        ['#0078d4','#2b88d8','#00b4d8','#038387','#007a7a','#0d73dd','#086f68','#00bcf2','#008272','#004e8c'],
};

// Scan multiple rows because REST APIs and XLSX.utils.sheet_to_json omit keys
// for missing values — a column absent from row 1 would otherwise never appear.
const COLUMN_SCAN_ROWS = 50;

export const extractColumns = (rows: IChartRecord[]): string[] => {
  const seen = new Set<string>();
  const excluded = new Set<string>();
  const columns: string[] = [];
  const scanCount = Math.min(rows.length, COLUMN_SCAN_ROWS);
  for (let i = 0; i < scanCount; i++) {
    // Guard against non-object entries (e.g. a null in a REST API array) —
    // Object.keys on those throws instead of just skipping the row.
    if (rows[i] === null || typeof rows[i] !== 'object') continue;
    for (const key of Object.keys(rows[i])) {
      // Object values (SharePoint lookup/person fields, nested JSON) can't be
      // charted or displayed — they'd render as "[object Object]".
      const value = rows[i][key] as unknown;
      if (value !== null && typeof value === 'object') {
        excluded.add(key);
        continue;
      }
      if (!seen.has(key) && !key.startsWith('odata.') && key !== '__metadata') {
        seen.add(key);
        columns.push(key);
      }
    }
  }
  return columns.filter(c => !excluded.has(c));
};

// A saved view state (Spotfire-style bookmark): data-shaping settings + column mapping
export interface IBookmark {
  name: string;
  state: {
    sortColumn: string;
    sortDirection: string;
    rowLimit: number;
    filterColumn: string;
    filterValue: string;
    groupByColumn: string;
    aggregation: string;
    xColumn: string;
    yColumns: string;
    // Optional so bookmarks saved before filter operators existed still load
    filterOperator?: string;
  };
}

export const parseBookmarks = (json: string): IBookmark[] => {
  if (!json) return [];
  try {
    const parsed = JSON.parse(json);
    if (!Array.isArray(parsed)) return [];
    // Reject malformed entries (hand-edited properties, imported page
    // templates) — applying one would throw deep in the apply handler and
    // unmount the whole web part for viewers.
    return parsed.filter((b): b is IBookmark =>
      !!b && typeof b === 'object' && typeof b.name === 'string' &&
      !!b.state && typeof b.state === 'object');
  } catch {
    return [];
  }
};

// ---- Row filters ----

export type FilterOperator =
  | 'contains' | 'equals' | 'notEquals' | 'gt' | 'lt' | 'between' | 'isEmpty' | 'notEmpty';

export const FILTER_OPERATORS: FilterOperator[] =
  ['contains', 'equals', 'notEquals', 'gt', 'lt', 'between', 'isEmpty', 'notEmpty'];

export const operatorNeedsValue = (op: string): boolean => op !== 'isEmpty' && op !== 'notEmpty';

export interface IRowFilter {
  column: string;
  operator: string;
  value: string;
}

const isBlank = (v: unknown): boolean => v === null || v === undefined || v === '';

// Numbers compare as numbers, parseable dates as dates, everything else as text
const compareCell = (cell: unknown, target: string): number => {
  const cn = Number(cell), tn = Number(target);
  if (!isBlank(cell) && target.trim() !== '' && !isNaN(cn) && !isNaN(tn)) return cn - tn;
  const cd = Date.parse(String(cell)), td = Date.parse(target);
  if (typeof cell === 'string' && isNaN(cn) && !isNaN(cd) && !isNaN(td)) return cd - td;
  return String(cell ?? '').toLowerCase().localeCompare(target.toLowerCase());
};

export const rowMatchesFilter = (row: IChartRecord, f: IRowFilter): boolean => {
  const cell = row[f.column];
  switch (f.operator || 'contains') {
    case 'isEmpty': return isBlank(cell);
    case 'notEmpty': return !isBlank(cell);
    case 'equals': return compareCell(cell, f.value) === 0;
    case 'notEquals': return compareCell(cell, f.value) !== 0;
    case 'gt': return !isBlank(cell) && compareCell(cell, f.value) > 0;
    case 'lt': return !isBlank(cell) && compareCell(cell, f.value) < 0;
    case 'between': {
      // "low..high" (inclusive)
      const parts = f.value.split('..');
      if (parts.length !== 2 || isBlank(cell)) return false;
      return compareCell(cell, parts[0].trim()) >= 0 && compareCell(cell, parts[1].trim()) <= 0;
    }
    default:
      return String(cell ?? '').toLowerCase().indexOf(f.value.toLowerCase()) >= 0;
  }
};

// A filter is active when it has a column and — for operators that take one — a value
export const isFilterActive = (f: IRowFilter): boolean =>
  !!f.column && (!operatorNeedsValue(f.operator || 'contains') || f.value !== '');

export const applyFilters = (rows: IChartRecord[], filters: IRowFilter[]): IChartRecord[] => {
  const active = filters.filter(isFilterActive);
  return active.length ? rows.filter(r => active.every(f => rowMatchesFilter(r, f))) : rows;
};

// Substitute {0}, {1}, … placeholders in localized string templates.
export const fmt = (template: string, ...args: (string | number)[]): string =>
  template.replace(/\{(\d+)\}/g, (match, idx) => {
    const arg = args[Number(idx)];
    return arg !== undefined ? String(arg) : match;
  });

// Prefix of the per-group error columns emitted when averages are aggregated with SD/SEM error bars
export const ERROR_COLUMN_PREFIX = '__err_';

// A user-entered color as an opaque 6-digit hex (safe to suffix with an alpha byte)
export const toSolidHex = (color: string, fallback: string): string => {
  const c = normalizeHexColor((color || '').trim());
  return /^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(c) ? c.slice(0, 7) : fallback;
};

// Expand a 3-digit hex ('#abc') to 6-digit ('#aabbcc'). ChartRenderer builds
// translucent fills by concatenating an alpha suffix directly onto these
// colors (e.g. `${color}cc`), which only produces a valid 8-digit hex color
// when the base is already 6 digits — a 3-digit override would silently
// produce an invalid color that canvas ignores.
export const normalizeHexColor = (color: string): string => {
  const m = /^#([0-9a-fA-F])([0-9a-fA-F])([0-9a-fA-F])$/.exec(color);
  return m ? `#${m[1]}${m[1]}${m[2]}${m[2]}${m[3]}${m[3]}` : color;
};

export const resolveColors = (palette: string, seriesColors: string, count: number): string[] => {
  const base = PALETTES[palette] || PALETTES.office;
  const overrides = seriesColors ? seriesColors.split(',') : [];
  return Array.from({ length: count }, (_, i) => {
    const override = overrides[i] ? overrides[i].trim() : '';
    return normalizeHexColor(override || base[i % base.length]);
  });
};

export const DATA_SOURCE_LABELS: Record<DataSourceType, string> = {
  upload: strings.SourceUploadLabel,
  sharePointList: strings.SourceSharePointListLabel,
  sharePointFile: strings.SourceSharePointFileLabel,
  restApi: strings.SourceRestApiLabel,
  graphApi: strings.SourceGraphApiLabel,
};

export const DATA_SOURCE_ICONS: Record<DataSourceType, string> = {
  upload: '📁',
  sharePointList: '📋',
  sharePointFile: '🔗',
  restApi: '🌐',
  graphApi: '🪐',
};

export const isPieOrDoughnut = (chartType: ChartType): boolean =>
  chartType === 'pie' || chartType === 'doughnut';

export const isScatterOrBubble = (chartType: ChartType): boolean =>
  chartType === 'scatter' || chartType === 'bubble';

export const needsNumericX = (chartType: ChartType): boolean =>
  chartType === 'scatter' || chartType === 'bubble';

// Chart types that take exactly one value (Y) column
export const isSingleValueType = (chartType: ChartType): boolean =>
  chartType === 'pie' || chartType === 'doughnut' || chartType === 'scatter' ||
  chartType === 'bubble' || chartType === 'kpi' || chartType === 'waterfall' ||
  chartType === 'boxplot' || chartType === 'violin' || chartType === 'treemap' ||
  chartType === 'heatmap';

// KPI shows a single aggregated number — no X axis at all
export const hasNoXColumn = (chartType: ChartType): boolean => chartType === 'kpi';

// Histogram bins its X column; no Y column needed
export const hasNoYColumn = (chartType: ChartType): boolean => chartType === 'histogram';

// Heatmap needs a second (row) category column
export const needsRowColumn = (chartType: ChartType): boolean => chartType === 'heatmap';

// Chart types that support error bars
export const supportsErrorBars = (chartType: ChartType): boolean =>
  chartType === 'bar' || chartType === 'horizontalBar' ||
  chartType === 'line' || chartType === 'area';

// Chart types that support a secondary (right) Y axis. horizontalBar is
// excluded — its value axis is x, so a second value axis would need a second
// x scale, not the y1 scale ChartRenderer implements.
export const supportsDualAxis = (chartType: ChartType): boolean =>
  chartType === 'bar' || chartType === 'line' || chartType === 'area';

// Chart types that support data point overlay
export const supportsDataPointOverlay = (chartType: ChartType): boolean =>
  chartType === 'bar' || chartType === 'horizontalBar';
