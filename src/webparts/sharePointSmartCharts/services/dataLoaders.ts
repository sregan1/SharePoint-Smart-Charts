import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import { IChartRecord, fmt } from '../types';

// SharePoint REST returns at most this many items per request; results at this
// count are likely truncated.
export const SP_LIST_ROW_LIMIT = 5000;

export interface ILoadResult {
  rows: IChartRecord[];
  truncated: boolean;
  /** Sheet names when the source was a multi-sheet Excel workbook */
  sheetNames?: string[];
  /** Rows PapaParse couldn't parse and skipped (not fatal — some rows still loaded) */
  parseWarningCount?: number;
}

export interface ICsvParseResult {
  rows: IChartRecord[];
  /** Rows PapaParse couldn't parse and skipped, e.g. a malformed row mid-file */
  errorCount: number;
}

export const parseCsvText = (text: string, delimiter?: string): ICsvParseResult => {
  const result = Papa.parse<IChartRecord>(text, {
    header: true,
    dynamicTyping: true,
    skipEmptyLines: true,
    delimiter: delimiter || undefined,
  });
  if (result.errors.length && !result.data.length) {
    throw new Error(result.errors[0].message);
  }
  return { rows: result.data, errorCount: result.errors.length };
};

const isMidnightUtc = (d: Date): boolean =>
  d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && d.getUTCMilliseconds() === 0;

export const sheetToRows = (workbook: XLSX.WorkBook, sheetName?: string): IChartRecord[] => {
  const name = sheetName && workbook.SheetNames.indexOf(sheetName) >= 0
    ? sheetName
    : workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json<IChartRecord>(workbook.Sheets[name]);
  // Reading the workbook with cellDates:true yields real Date objects for date
  // cells (instead of meaningless serial numbers) — convert to ISO strings so
  // they survive persistence and extractColumns, which excludes object values.
  return rows.map(row => {
    const out: IChartRecord = { ...row };
    for (const key of Object.keys(out)) {
      const v = out[key] as unknown;
      if (v instanceof Date) {
        out[key] = isMidnightUtc(v) ? v.toISOString().slice(0, 10) : v.toISOString();
      }
    }
    return out;
  });
};

export const parseExcelBuffer = (buffer: ArrayBuffer, sheetName?: string): IChartRecord[] =>
  sheetToRows(XLSX.read(buffer, { type: 'array', cellDates: true }), sheetName);

const isExcelUrl = (url: string): boolean => {
  const path = url.split('?')[0].toLowerCase();
  return path.endsWith('.xlsx') || path.endsWith('.xls') || path.endsWith('.xlsm') || path.endsWith('.xlsb');
};

// Plumbing fields select('*') always returns that are never meaningful to
// chart — excluded so they don't flood the column dropdowns. Deliberately
// conservative: genuinely useful fields (ID, Created, Modified, Author) stay.
const SP_INTERNAL_FIELDS = new Set([
  'ContentTypeId',
  'FileSystemObjectType',
  'GUID',
  'OData__UIVersionString',
  'OData__ColorTag',
  'ComplianceAssetId',
  'ServerRedirectedEmbedUri',
  'ServerRedirectedEmbedUrl',
]);

export const loadSharePointList = async (
  context: WebPartContext,
  siteUrl: string,
  listName: string
): Promise<ILoadResult> => {
  const sp = spfi(siteUrl || context.pageContext.web.absoluteUrl).using(SPFx(context));
  const rawRows = (await sp.web.lists
    .getByTitle(listName)
    .items.select('*')
    .top(SP_LIST_ROW_LIMIT)()) as IChartRecord[];
  const rows = rawRows.map(row => {
    const out: IChartRecord = {};
    for (const key of Object.keys(row)) {
      if (!SP_INTERNAL_FIELDS.has(key)) out[key] = row[key];
    }
    return out;
  });
  return { rows, truncated: rows.length >= SP_LIST_ROW_LIMIT };
};

const FETCH_TIMEOUT_MS = 30_000;

// A hung endpoint would otherwise pin the data source panel in "Loading..."
// forever with no feedback — abort and surface a clear error instead.
const fetchWithTimeout = async (url: string, init?: RequestInit): Promise<Response> => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } catch (e) {
    if (e instanceof Error && e.name === 'AbortError') {
      throw new Error(strings.ErrorRequestTimedOut);
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }
};

export const loadSharePointFile = async (
  dataUrl: string,
  delimiter?: string,
  sheetName?: string
): Promise<ILoadResult> => {
  const response = await fetchWithTimeout(dataUrl, { credentials: 'same-origin' });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  if (isExcelUrl(dataUrl)) {
    const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array', cellDates: true });
    return {
      rows: sheetToRows(workbook, sheetName),
      truncated: false,
      sheetNames: workbook.SheetNames.slice(),
    };
  }
  const text = await response.text();
  const contentType = response.headers.get('content-type') || '';
  const looksLikeHtml = contentType.indexOf('text/html') >= 0 || /^\s*<(!doctype|html)/i.test(text);
  if (looksLikeHtml) {
    throw new Error(strings.ErrorUrlReturnedHtml);
  }
  const parsed = parseCsvText(text, delimiter);
  return { rows: parsed.rows, truncated: false, parseWarningCount: parsed.errorCount };
};

const applyDataPath = (json: any, dataPath?: string): IChartRecord[] => {
  let result = json;
  if (dataPath) {
    for (const part of dataPath.split('.')) result = result?.[part];
    if (result === undefined || result === null) {
      throw new Error(fmt(strings.ErrorDataPathNotFound, dataPath));
    }
  } else if (!Array.isArray(result) && Array.isArray(result?.value)) {
    // OData-style envelope (SharePoint REST, Microsoft Graph) — unwrap automatically
    result = result.value;
  }
  return Array.isArray(result) ? result : [result];
};

export const loadRestApi = async (
  dataUrl: string,
  dataPath?: string
): Promise<ILoadResult> => {
  const response = await fetchWithTimeout(dataUrl, {
    headers: { Accept: 'application/json' },
    credentials: 'same-origin',
  });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  let json: unknown;
  try {
    json = await response.json();
  } catch {
    // Don't rely on the content-type header — many APIs omit or misreport it.
    // A real parse failure (HTML error page with a 200 status, malformed
    // body) is a clearer message than the raw "Unexpected token <" SyntaxError.
    throw new Error(strings.ErrorNonJsonResponse);
  }
  return { rows: applyDataPath(json, dataPath), truncated: false };
};

// Microsoft Graph pages most collections at ~100 items via @odata.nextLink.
// A single .get() silently returns only the first page — follow nextLink up
// to this many pages (matches SP_LIST_ROW_LIMIT at ~100/page) before giving up
// and flagging the result as truncated, same as the SharePoint list loader.
export const GRAPH_MAX_PAGES = 50;

export const loadGraphApi = async (
  context: WebPartContext,
  graphPath: string,
  dataPath?: string
): Promise<ILoadResult> => {
  const client = await context.msGraphClientFactory.getClient('3');
  let rows: IChartRecord[] = [];
  let nextUrl: string | undefined = graphPath;
  let pages = 0;

  while (nextUrl && pages < GRAPH_MAX_PAGES) {
    const json = await client.api(nextUrl).get();
    rows = rows.concat(applyDataPath(json, dataPath));
    nextUrl = json?.['@odata.nextLink'];
    pages++;
  }

  return { rows, truncated: !!nextUrl };
};

// ---- Session cache (REST/Graph results) ----

// Single source of truth for the cache key shape — used by the auto-load
// effect, the manual "Refresh" button, and DataSourcePanel's manual load
// handlers, so they can never drift apart and target different entries.
export const buildCacheKey = (srcType: string, dataUrl: string, dataPath?: string): string =>
  `${srcType}|${dataUrl}|${dataPath || ''}`;

const CACHE_PREFIX = 'sdv-cache:';

export const getCachedRows = (key: string, maxAgeMinutes: number): IChartRecord[] | null => {
  if (maxAgeMinutes <= 0) return null;
  try {
    const raw = sessionStorage.getItem(CACHE_PREFIX + key);
    if (!raw) return null;
    const entry = JSON.parse(raw) as { ts: number; rows: IChartRecord[] };
    if (Date.now() - entry.ts > maxAgeMinutes * 60_000) return null;
    return entry.rows;
  } catch {
    return null;
  }
};

export const setCachedRows = (key: string, rows: IChartRecord[]): void => {
  try {
    sessionStorage.setItem(CACHE_PREFIX + key, JSON.stringify({ ts: Date.now(), rows }));
  } catch { /* storage full or unavailable — caching is best-effort */ }
};

export const clearCachedRows = (key: string): void => {
  try {
    sessionStorage.removeItem(CACHE_PREFIX + key);
  } catch { /* ignore */ }
};
