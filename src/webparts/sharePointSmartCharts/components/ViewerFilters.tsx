import * as React from 'react';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import {
  IChartRecord,
  IRowFilter,
  FILTER_OPERATORS,
  operatorNeedsValue,
} from '../types';
import styles from './SharePointSmartCharts.module.scss';

interface IViewerFiltersProps {
  columns: string[];
  data: IChartRecord[];
  filters: IRowFilter[];
  onChange: (filters: IRowFilter[]) => void;
}

const MAX_DISTINCT_VALUES = 50;

export const FILTER_OPERATOR_LABELS: Record<string, string> = {
  contains: strings.FilterOpContains,
  equals: strings.FilterOpEquals,
  notEquals: strings.FilterOpNotEquals,
  gt: strings.FilterOpGt,
  lt: strings.FilterOpLt,
  between: strings.FilterOpBetween,
  isEmpty: strings.FilterOpIsEmpty,
  notEmpty: strings.FilterOpNotEmpty,
};

// Distinct values of a column, or undefined when there are too many to list
const distinctValues = (data: IChartRecord[], column: string): string[] | undefined => {
  const seen = new Set<string>();
  for (const row of data) {
    const v = row[column];
    if (v === null || v === undefined || v === '') continue;
    seen.add(String(v));
    if (seen.size > MAX_DISTINCT_VALUES) return undefined;
  }
  return Array.from(seen).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
};

const ViewerFilters: React.FC<IViewerFiltersProps> = ({ columns, data, filters, onChange }) => {
  const rows: IRowFilter[] = filters.length ? filters : [{ column: '', operator: 'contains', value: '' }];

  const update = (index: number, partial: Partial<IRowFilter>) =>
    onChange(rows.map((f, i) => (i === index ? { ...f, ...partial } : f)));

  const remove = (index: number) => {
    const next = rows.filter((_, i) => i !== index);
    onChange(next);
  };

  return (
    <div className={styles.viewerFilterStack}>
      {rows.map((f, i) => {
        const options = f.column && (f.operator === 'equals' || f.operator === 'notEquals')
          ? distinctValues(data, f.column)
          : undefined;
        const needsValue = operatorNeedsValue(f.operator);
        return (
          <div className={styles.viewerFilterBar} key={i}>
            {i === 0 && <span className={styles.viewerFilterLabel}>{strings.ViewerFilterLabel}</span>}
            <select
              value={f.column}
              onChange={e => update(i, { column: e.target.value, value: '' })}
              aria-label={strings.ViewerFilterColumnAria}
            >
              <option value="">{strings.NoneOption}</option>
              {columns.map(col => <option key={col} value={col}>{col}</option>)}
            </select>
            <select
              value={f.operator || 'contains'}
              onChange={e => update(i, { operator: e.target.value, value: '' })}
              disabled={!f.column}
              aria-label={strings.ViewerFilterOperatorAria}
            >
              {FILTER_OPERATORS.map(op => (
                <option key={op} value={op}>{FILTER_OPERATOR_LABELS[op]}</option>
              ))}
            </select>
            {needsValue && (options ? (
              <select
                value={f.value}
                onChange={e => update(i, { value: e.target.value })}
                disabled={!f.column}
                aria-label={strings.ViewerFilterValueAria}
              >
                <option value="">{strings.SelectValuePlaceholder}</option>
                {options.map(v => <option key={v} value={v}>{v}</option>)}
              </select>
            ) : (
              <input
                type="text"
                value={f.value}
                onChange={e => update(i, { value: e.target.value })}
                placeholder={f.operator === 'between' ? '10..20' : strings.FilterValuePlaceholder}
                disabled={!f.column}
                aria-label={strings.ViewerFilterValueAria}
              />
            ))}
            {(f.column || f.value || rows.length > 1) && (
              <button className={styles.secondaryButton} onClick={() => remove(i)}>
                {rows.length > 1 ? strings.RemoveFilterButton : strings.ClearButton}
              </button>
            )}
            {i === rows.length - 1 && f.column && (
              <button
                className={styles.secondaryButton}
                onClick={() => onChange([...rows, { column: '', operator: 'contains', value: '' }])}
              >
                {strings.AddFilterButton}
              </button>
            )}
          </div>
        );
      })}
    </div>
  );
};

export default ViewerFilters;
