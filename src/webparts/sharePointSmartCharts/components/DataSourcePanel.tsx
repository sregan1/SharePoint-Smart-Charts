import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import {
  IDataSourceConfig,
  IChartRecord,
  DataSourceType,
  DATA_SOURCE_LABELS,
  DATA_SOURCE_ICONS,
  extractColumns,
  fmt,
} from '../types';
import * as XLSX from 'xlsx';
import {
  parseCsvText,
  sheetToRows,
  loadSharePointList,
  loadSharePointFile,
  loadRestApi,
  loadGraphApi,
  SP_LIST_ROW_LIMIT,
  GRAPH_MAX_PAGES,
  clearCachedRows,
  buildCacheKey,
} from '../services/dataLoaders';
import styles from './SharePointSmartCharts.module.scss';

const SIZE_LIMIT = 200_000; // ~200KB JSON — safe SPFx property bag limit

interface IDataSourcePanelProps {
  config: IDataSourceConfig;
  context: WebPartContext;
  uploadedFileName: string;
  uploadedRowCount: number;
  onConfigChange: (config: Partial<IDataSourceConfig>) => void;
  onDataLoaded: (data: IChartRecord[], columns: string[]) => void;
  onPersistData: (json: string, fileName: string) => void;
  onClearData: () => void;
}

const DELIMITER_OPTIONS = [
  { value: '', label: strings.DelimiterAutoDetect },
  { value: ',', label: strings.DelimiterComma },
  { value: '\t', label: strings.DelimiterTab },
  { value: ';', label: strings.DelimiterSemicolon },
  { value: '|', label: strings.DelimiterPipe },
];

const DataSourcePanel: React.FC<IDataSourcePanelProps> = ({
  config,
  context,
  uploadedFileName,
  uploadedRowCount,
  onConfigChange,
  onDataLoaded,
  onPersistData,
  onClearData,
}) => {
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [success, setSuccess] = React.useState('');
  const [warning, setWarning] = React.useState('');
  const [availableLists, setAvailableLists] = React.useState<string[]>([]);
  const [isDiscovering, setIsDiscovering] = React.useState(false);
  const [discoverError, setDiscoverError] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  // Kept so the user can switch sheets without re-uploading the file
  const workbookRef = React.useRef<XLSX.WorkBook | null>(null);
  // Kept so the delimiter dropdown can re-parse an uploaded text file without
  // requiring a re-upload (the Load button is hidden once a file is loaded).
  const rawTextRef = React.useRef<string | null>(null);
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  // Unique id prefix so label/input pairing stays valid with multiple instances on a page
  const idPrefix = React.useRef(`sdv-${Math.random().toString(36).slice(2, 8)}`).current;

  const sourceTypes: DataSourceType[] = ['upload', 'sharePointList', 'sharePointFile', 'restApi', 'graphApi'];

  React.useEffect(() => {
    if (config.dataSourceType !== 'sharePointList') return;
    let cancelled = false;
    // Debounce — config.siteUrl updates on every keystroke, and firing a PnPjs
    // request per keystroke against a partially-typed URL risks 429 throttling.
    const timer = setTimeout(() => {
      const discover = async () => {
        setIsDiscovering(true);
        setDiscoverError('');
        try {
          const sp = spfi(config.siteUrl || context.pageContext.web.absoluteUrl).using(SPFx(context));
          const lists = await sp.web.lists
            .filter('Hidden eq false')
            .select('Title')
            .orderBy('Title', true)();
          if (!cancelled) {
            setAvailableLists((lists as Array<{ Title: string }>).map(l => l.Title));
          }
        } catch {
          if (!cancelled) setDiscoverError(strings.ListDiscoveryError);
        } finally {
          if (!cancelled) setIsDiscovering(false);
        }
      };
      discover();
    }, 500);
    return () => { cancelled = true; clearTimeout(timer); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [config.dataSourceType, config.siteUrl]);

  const handleDataLoaded = (data: IChartRecord[], fileName?: string, parseWarningCount?: number) => {
    const columns = extractColumns(data);
    setError('');
    setWarning('');

    if (fileName !== undefined) {
      // File upload — attempt persistence
      const json = JSON.stringify(data);
      if (json.length <= SIZE_LIMIT) {
        onPersistData(json, fileName);
        setSuccess(fmt(strings.LoadedRowsColumnsPersist, data.length, columns.length));
      } else {
        onPersistData('', fileName); // store filename only so UI reflects "loaded" state
        setWarning(fmt(strings.SizeWarning, Math.round(json.length / 1024)));
        setSuccess(fmt(strings.LoadedRowsColumns, data.length, columns.length));
      }
    } else {
      setSuccess(fmt(strings.LoadedRowsColumns, data.length, columns.length));
    }

    // A partial CSV parse (some rows malformed) isn't fatal — the rest still
    // loaded — but the user should know rows were silently dropped.
    if (parseWarningCount) {
      setWarning(fmt(strings.CsvParseWarning, parseWarningCount));
    }

    onDataLoaded(data, columns);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    // Reset input so the same file can be re-selected if needed
    e.target.value = '';
    setLoading(true);
    setError('');
    setSuccess('');
    setWarning('');
    try {
      const name = file.name.toLowerCase();
      let data: IChartRecord[];
      let parseWarningCount = 0;
      if (name.endsWith('.csv') || name.endsWith('.tsv') || name.endsWith('.txt')) {
        const text = await file.text();
        rawTextRef.current = text;
        const parsed = parseCsvText(text, config.delimiter || undefined);
        data = parsed.rows;
        parseWarningCount = parsed.errorCount;
        workbookRef.current = null;
        setSheetNames([]);
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.xlsm') || name.endsWith('.xlsb')) {
        rawTextRef.current = null;
        const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
        workbookRef.current = workbook;
        setSheetNames(workbook.SheetNames.length > 1 ? workbook.SheetNames.slice() : []);
        data = sheetToRows(workbook, config.sheetName || undefined);
      } else {
        throw new Error(strings.ErrorUnsupportedFileType);
      }
      handleDataLoaded(data, file.name, parseWarningCount);
    } catch (err) {
      setError(err instanceof Error ? err.message : strings.ErrorParseFile);
    } finally {
      setLoading(false);
    }
  };

  const handleSharePointList = async () => {
    if (!config.listName) { setError(strings.ErrorSelectListName); return; }
    setLoading(true);
    setError('');
    try {
      const result = await loadSharePointList(context, config.siteUrl, config.listName);
      handleDataLoaded(result.rows);
      if (result.truncated) {
        setWarning(fmt(strings.ListTruncatedWarning, SP_LIST_ROW_LIMIT.toLocaleString()));
      }
    } catch (err) {
      setError(fmt(strings.ErrorLoadList, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleSharePointFile = async (sheetOverride?: string) => {
    if (!config.dataUrl) { setError(strings.ErrorEnterFileUrl); return; }
    setLoading(true);
    setError('');
    try {
      const sheet = sheetOverride !== undefined ? sheetOverride : (config.sheetName || undefined);
      const result = await loadSharePointFile(config.dataUrl, config.delimiter || undefined, sheet);
      setSheetNames(result.sheetNames && result.sheetNames.length > 1 ? result.sheetNames : []);
      handleDataLoaded(result.rows, undefined, result.parseWarningCount);
    } catch (err) {
      setError(fmt(strings.ErrorLoadFile, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleRestApi = async () => {
    if (!config.dataUrl) { setError(strings.ErrorEnterUrl); return; }
    setLoading(true);
    setError('');
    try {
      const result = await loadRestApi(config.dataUrl, config.dataPath || undefined);
      handleDataLoaded(result.rows);
      // A manual load reflects the latest server data — clear any older
      // cached entry so a later page reload doesn't resurrect stale rows.
      clearCachedRows(buildCacheKey('restApi', config.dataUrl, config.dataPath));
    } catch (err) {
      setError(fmt(strings.ErrorFetchData, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleGraphApi = async () => {
    if (!config.dataUrl) { setError(strings.ErrorEnterUrl); return; }
    setLoading(true);
    setError('');
    try {
      const result = await loadGraphApi(context, config.dataUrl, config.dataPath || undefined);
      handleDataLoaded(result.rows);
      clearCachedRows(buildCacheKey('graphApi', config.dataUrl, config.dataPath));
      if (result.truncated) {
        setWarning(fmt(strings.GraphTruncatedWarning, GRAPH_MAX_PAGES));
      }
    } catch (err) {
      setError(fmt(strings.ErrorFetchData, err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const handleLoad = () => {
    switch (config.dataSourceType) {
      case 'upload': fileInputRef.current?.click(); break;
      case 'sharePointList': handleSharePointList(); break;
      case 'sharePointFile': handleSharePointFile(); break;
      case 'restApi': handleRestApi(); break;
      case 'graphApi': handleGraphApi(); break;
    }
  };

  const handleDelimiterChange = (delimiter: string) => {
    onConfigChange({ delimiter });
    // The Load button is hidden once an upload has loaded — re-parse the
    // retained raw text so changing the delimiter actually takes effect.
    if (config.dataSourceType === 'upload' && rawTextRef.current !== null) {
      try {
        const parsed = parseCsvText(rawTextRef.current, delimiter || undefined);
        handleDataLoaded(parsed.rows, uploadedFileName || config.uploadedFileName, parsed.errorCount);
      } catch (err) {
        setError(err instanceof Error ? err.message : strings.ErrorParseFile);
      }
    }
  };

  const handleSheetChange = (name: string) => {
    onConfigChange({ sheetName: name });
    if (config.dataSourceType === 'upload' && workbookRef.current) {
      const rows = sheetToRows(workbookRef.current, name);
      handleDataLoaded(rows, uploadedFileName || config.uploadedFileName);
    } else if (config.dataSourceType === 'sharePointFile') {
      handleSharePointFile(name);
    }
  };

  const handleChangeFile = () => {
    onClearData();
    setSuccess('');
    setWarning('');
    workbookRef.current = null;
    rawTextRef.current = null;
    setSheetNames([]);
    // Small delay so state clears before the picker opens
    setTimeout(() => fileInputRef.current?.click(), 50);
  };

  const isTextFile = (name: string) => /\.(csv|tsv|txt)(\?.*)?$/i.test(name);
  const isUpload = config.dataSourceType === 'upload';
  const hasLoadedFile = isUpload && !!uploadedFileName;
  const showDelimiter =
    (isUpload && !!uploadedFileName && isTextFile(uploadedFileName)) ||
    (config.dataSourceType === 'sharePointFile' && !!config.dataUrl && isTextFile(config.dataUrl));

  return (
    <div className={styles.editPanel}>
      <div className={styles.sectionHeader}>{strings.DataSourceSectionHeader}</div>

      <div className={styles.sourceTypeGrid}>
        {sourceTypes.map(type => (
          <button
            key={type}
            className={`${styles.sourceTypeCard} ${config.dataSourceType === type ? styles.selected : ''}`}
            onClick={() => {
              onConfigChange({ dataSourceType: type });
              setError(''); setSuccess(''); setWarning('');
              // A sheet picker or delimiter left over from a previous source
              // type is meaningless once the source type has changed.
              workbookRef.current = null; rawTextRef.current = null; setSheetNames([]);
            }}
            aria-pressed={config.dataSourceType === type}
          >
            <span className={styles.icon} aria-hidden="true">{DATA_SOURCE_ICONS[type]}</span>
            <span className={styles.label}>{DATA_SOURCE_LABELS[type]}</span>
          </button>
        ))}
      </div>

      {/* Upload — two-state UI */}
      {isUpload && (
        <div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.tsv,.txt,.xlsx,.xls,.xlsm,.xlsb"
            style={{ display: 'none' }}
            onChange={handleFileUpload}
            aria-label={strings.UploadHelp}
          />
          {hasLoadedFile ? (
            <div className={styles.loadedFileBanner}>
              <span className={styles.loadedFileIcon} aria-hidden="true">📁</span>
              <span className={styles.loadedFileInfo}>
                <strong>{uploadedFileName}</strong>
                {uploadedRowCount > 0 && (
                  <span className={styles.loadedFileRows}> {fmt(strings.LoadedFileRowsLabel, uploadedRowCount.toLocaleString())}</span>
                )}
              </span>
              <div className={styles.loadedFileActions}>
                <button className={styles.secondaryButton} onClick={handleChangeFile} disabled={loading}>
                  {strings.ChangeFileButton}
                </button>
                <button
                  className={styles.secondaryButton}
                  onClick={() => { onClearData(); setSuccess(''); setWarning(''); workbookRef.current = null; rawTextRef.current = null; setSheetNames([]); }}
                  disabled={loading}
                >
                  {strings.ClearButton}
                </button>
              </div>
            </div>
          ) : (
            <p className={styles.helpText}>{strings.UploadHelp}</p>
          )}
        </div>
      )}

      {(showDelimiter || sheetNames.length > 1) && (
        <div className={styles.fieldRow}>
          {showDelimiter && (
            <div className={styles.fieldGroup} style={{ flex: '0 0 auto', width: 160 }}>
              <label htmlFor={`${idPrefix}-delimiter`}>{strings.DelimiterLabel}</label>
              <select
                id={`${idPrefix}-delimiter`}
                value={config.delimiter || ''}
                onChange={e => handleDelimiterChange(e.target.value)}
              >
                {DELIMITER_OPTIONS.map(opt => (
                  <option key={opt.label} value={opt.value}>{opt.label}</option>
                ))}
              </select>
            </div>
          )}
          {sheetNames.length > 1 && (
            <div className={styles.fieldGroup}>
              <label htmlFor={`${idPrefix}-sheet`}>{strings.SheetNameLabel}</label>
              <select
                id={`${idPrefix}-sheet`}
                value={config.sheetName && sheetNames.indexOf(config.sheetName) >= 0 ? config.sheetName : sheetNames[0]}
                onChange={e => handleSheetChange(e.target.value)}
                disabled={loading}
              >
                {sheetNames.map(name => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
            </div>
          )}
        </div>
      )}

      {config.dataSourceType === 'sharePointList' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-siteurl`}>{strings.SiteUrlLabel}</label>
            <input
              id={`${idPrefix}-siteurl`}
              type="url"
              value={config.siteUrl}
              onChange={e => onConfigChange({ siteUrl: e.target.value })}
              placeholder={context.pageContext.web.absoluteUrl}
            />
            <span className={styles.helpText}>{strings.SiteUrlHelp}</span>
          </div>
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-listname`}>
              {strings.ListNameLabel}
              {isDiscovering && <span className={styles.helpText}> {strings.LoadingListsLabel}</span>}
            </label>
            {availableLists.length > 0 ? (
              <select
                id={`${idPrefix}-listname`}
                value={config.listName}
                onChange={e => onConfigChange({ listName: e.target.value })}
              >
                <option value="">{strings.SelectListPlaceholder}</option>
                {availableLists.map(name => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
            ) : (
              <input
                id={`${idPrefix}-listname`}
                type="text"
                value={config.listName}
                onChange={e => onConfigChange({ listName: e.target.value })}
                placeholder={discoverError ? strings.EnterListNamePlaceholder : strings.ListNameExamplePlaceholder}
              />
            )}
            {discoverError && <span className={styles.helpText}>{discoverError}</span>}
          </div>
        </div>
      )}

      {config.dataSourceType === 'sharePointFile' && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
            <label htmlFor={`${idPrefix}-fileurl`}>{strings.FileUrlLabel}</label>
            <input
              id={`${idPrefix}-fileurl`}
              type="url"
              value={config.dataUrl}
              onChange={e => onConfigChange({ dataUrl: e.target.value })}
              placeholder="https://yourtenant.sharepoint.com/sites/mysite/Shared Documents/data.csv"
            />
            <span className={styles.helpText}>{strings.FileUrlHelp}</span>
          </div>
        </div>
      )}

      {(config.dataSourceType === 'restApi' || config.dataSourceType === 'graphApi') && (
        <div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
              <label htmlFor={`${idPrefix}-apiurl`}>
                {config.dataSourceType === 'graphApi' ? strings.GraphPathLabel : strings.ApiUrlLabel}
              </label>
              <input
                id={`${idPrefix}-apiurl`}
                type={config.dataSourceType === 'graphApi' ? 'text' : 'url'}
                value={config.dataUrl}
                onChange={e => onConfigChange({ dataUrl: e.target.value })}
                placeholder={config.dataSourceType === 'graphApi' ? '/me/memberOf' : 'https://api.example.com/data'}
              />
              {config.dataSourceType === 'graphApi' && (
                <span className={styles.helpText}>{strings.GraphPathHelp}</span>
              )}
            </div>
          </div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup}>
              <label htmlFor={`${idPrefix}-datapath`}>{strings.DataPathLabel}</label>
              <input
                id={`${idPrefix}-datapath`}
                type="text"
                value={config.dataPath}
                onChange={e => onConfigChange({ dataPath: e.target.value })}
                placeholder="e.g. value  or  data.items"
              />
              <span className={styles.helpText}>{strings.DataPathHelp}</span>
            </div>
          </div>
        </div>
      )}

      {warning && (
        <div className={styles.warningMessage} role="alert">⚠️ {warning}</div>
      )}
      {error && <div className={styles.errorMessage} role="alert">⚠️ {error}</div>}
      {success && <div className={styles.successMessage} role="status">✓ {success}</div>}

      {/* Hide Load button for upload when a file is already loaded (use Change File instead) */}
      {!(isUpload && hasLoadedFile) && (
        <div className={styles.buttonRow}>
          <button
            className={styles.primaryButton}
            onClick={handleLoad}
            disabled={loading}
          >
            {loading ? strings.LoadingLabel : config.dataSourceType === 'upload' ? strings.ChooseFileButton : strings.LoadDataButton}
          </button>
        </div>
      )}
    </div>
  );
};

export default DataSourcePanel;
