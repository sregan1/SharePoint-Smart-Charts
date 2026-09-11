import * as React from 'react';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import {
  ChartType,
  IColumnConfig,
  PALETTES,
  isPieOrDoughnut,
  isScatterOrBubble,
  isSingleValueType,
  hasNoXColumn,
  hasNoYColumn,
  needsRowColumn,
  fmt,
} from '../types';
import styles from './SharePointSmartCharts.module.scss';

interface IColumnMapperProps {
  columns: string[];
  numericColumns?: string[];
  config: IColumnConfig;
  chartType: ChartType;
  colorPalette: string;
  seriesColors: string;
  seriesTypes: string;
  showAdvanced: boolean;
  onChange: (config: Partial<IColumnConfig>) => void;
  onSeriesColorsChange: (colors: string) => void;
  onSeriesTypesChange: (types: string) => void;
}

const ColumnMapper: React.FC<IColumnMapperProps> = ({
  columns,
  numericColumns,
  config,
  chartType,
  colorPalette,
  seriesColors,
  seriesTypes,
  showAdvanced,
  onChange,
  onSeriesColorsChange,
  onSeriesTypesChange,
}) => {
  const isRadar = chartType === 'radar';
  const isPie = isPieOrDoughnut(chartType);
  const isScatterBubble = isScatterOrBubble(chartType);
  const isBubble = chartType === 'bubble';
  // Unique id prefix so label/input pairing stays valid with multiple instances on a page
  const idPrefix = React.useRef(`sdv-cm-${Math.random().toString(36).slice(2, 8)}`).current;

  const colorOverrides = seriesColors ? seriesColors.split(',') : [];
  const typeOverrides = seriesTypes ? seriesTypes.split(',') : [];
  // Per-series type only makes sense for cartesian multi-series charts (advanced mode only)
  const supportsComboTypes = showAdvanced &&
    (chartType === 'bar' || chartType === 'line' || chartType === 'area');

  const getSeriesColor = (yColIndex: number): string => {
    const c = colorOverrides[yColIndex] ? colorOverrides[yColIndex].trim() : '';
    if (c) return c;
    // Match the palette color the chart actually renders for this series
    // (ChartRenderer's resolveColors) instead of a fixed fallback blue.
    const base = PALETTES[colorPalette] || PALETTES.office;
    return base[yColIndex % base.length];
  };

  const setSeriesColor = (yColIndex: number, hex: string) => {
    const current = [...colorOverrides];
    while (current.length <= yColIndex) current.push('');
    current[yColIndex] = hex;
    onSeriesColorsChange(current.join(','));
  };

  const getSeriesType = (yColIndex: number): string =>
    typeOverrides[yColIndex] ? typeOverrides[yColIndex].trim() : '';

  const setSeriesType = (yColIndex: number, type: string) => {
    const current = [...typeOverrides];
    while (current.length <= yColIndex) current.push('');
    current[yColIndex] = type;
    onSeriesTypesChange(current.join(','));
  };

  const toggleYColumn = (col: string) => {
    const current = config.yColumns || [];
    const removeIdx = current.indexOf(col);
    if (removeIdx >= 0) {
      // Removing: splice the same position out of the per-series color/type
      // overrides too, so later columns don't inherit the removed column's
      // color/type once yColumns shifts down.
      onChange({ yColumns: current.filter(c => c !== col) });
      if (colorOverrides.length > removeIdx) {
        const nextColors = [...colorOverrides];
        nextColors.splice(removeIdx, 1);
        onSeriesColorsChange(nextColors.join(','));
      }
      if (typeOverrides.length > removeIdx) {
        const nextTypes = [...typeOverrides];
        nextTypes.splice(removeIdx, 1);
        onSeriesTypesChange(nextTypes.join(','));
      }
    } else {
      onChange({ yColumns: [...current, col] });
    }
  };

  const showX = !hasNoXColumn(chartType);
  const showY = !hasNoYColumn(chartType);
  const showRowColumn = needsRowColumn(chartType);
  const singleY = isSingleValueType(chartType);

  const xLabel =
    isRadar ? strings.CategoryColumnLabel :
    isPie ? strings.LabelColumnLabel :
    chartType === 'histogram' ? strings.HistogramColumnLabel :
    chartType === 'treemap' ? strings.TreemapGroupLabel :
    chartType === 'heatmap' ? strings.HeatmapColumnLabel :
    chartType === 'waterfall' || chartType === 'boxplot' ? strings.CategoryColumnLabel :
    strings.XAxisColumnLabel;
  const yLabel =
    isScatterBubble ? strings.YAxisColumnLabel :
    singleY ? strings.ValueColumnLabel :
    strings.YAxisColumnsLabel;

  return (
    <div className={styles.editPanel}>
      <div className={styles.sectionHeader}>{strings.ColumnMappingSectionHeader}</div>
      <div className={styles.fieldRow}>
        {showX && (
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-x`}>{xLabel}</label>
            <select
              id={`${idPrefix}-x`}
              value={config.xColumn}
              onChange={e => onChange({ xColumn: e.target.value })}
            >
              <option value="">{strings.SelectColumnPlaceholder}</option>
              {(isScatterBubble && numericColumns ? numericColumns : columns).map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}

        {isScatterBubble && (
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-label`}>{strings.ScatterLabelColumnLabel}</label>
            <select
              id={`${idPrefix}-label`}
              value={config.labelColumn}
              onChange={e => onChange({ labelColumn: e.target.value })}
            >
              <option value="">{strings.NoSizeColumnOption}</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}

        {showRowColumn && (
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-row`}>{strings.HeatmapRowLabel}</label>
            <select
              id={`${idPrefix}-row`}
              value={config.labelColumn}
              onChange={e => onChange({ labelColumn: e.target.value })}
            >
              <option value="">{strings.SelectColumnPlaceholder}</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}

        {singleY && showY && (
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-y`}>{yLabel}</label>
            <select
              id={`${idPrefix}-y`}
              value={config.yColumns?.[0] || ''}
              onChange={e => onChange({ yColumns: [e.target.value] })}
            >
              <option value="">{strings.SelectColumnPlaceholder}</option>
              {(isScatterBubble && numericColumns ? numericColumns : columns).map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}

        {isBubble && (
          <div className={styles.fieldGroup}>
            <label htmlFor={`${idPrefix}-size`}>{strings.SizeColumnLabel}</label>
            <select
              id={`${idPrefix}-size`}
              value={config.sizeColumn}
              onChange={e => onChange({ sizeColumn: e.target.value })}
            >
              <option value="">{strings.NoSizeColumnOption}</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        )}
      </div>

      {!singleY && showY && (
        <div className={styles.fieldRow}>
          <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
            <label>{yLabel} <span className={styles.helpText}>{strings.YAxisColumnsHelp}</span></label>
            <ul className={styles.multiSelect}>
              {columns.map((col, colIdx) => {
                const checked = config.yColumns?.includes(col) || false;
                const yIdx = (config.yColumns || []).indexOf(col);
                return (
                  <li
                    key={col}
                    className={checked ? styles.selected : ''}
                  >
                    <label>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleYColumn(col)}
                      />
                      {col}
                    </label>
                    {checked && yIdx >= 0 && (
                      <>
                        <input
                          type="color"
                          className={styles.colorSwatch}
                          value={getSeriesColor(yIdx)}
                          onChange={e => setSeriesColor(yIdx, e.target.value)}
                          title={fmt(strings.SeriesColorTitle, col)}
                          aria-label={fmt(strings.SeriesColorTitle, col)}
                        />
                        {supportsComboTypes && (
                          <select
                            value={getSeriesType(yIdx)}
                            onChange={e => setSeriesType(yIdx, e.target.value)}
                            title={fmt(strings.SeriesTypeTitle, col)}
                            aria-label={fmt(strings.SeriesTypeTitle, col)}
                          >
                            <option value="">{strings.SeriesTypeDefault}</option>
                            <option value="bar">{strings.SeriesTypeBar}</option>
                            <option value="line">{strings.SeriesTypeLine}</option>
                          </select>
                        )}
                      </>
                    )}
                  </li>
                );
              })}
            </ul>
          </div>
        </div>
      )}
    </div>
  );
};

export default ColumnMapper;
