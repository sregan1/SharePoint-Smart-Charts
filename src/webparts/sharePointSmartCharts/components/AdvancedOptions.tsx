import * as React from 'react';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import { ChartType, IBookmark, isScatterOrBubble, fmt } from '../types';
import styles from './SharePointSmartCharts.module.scss';

interface IAdvancedOptionsProps {
  columns: string[];
  chartType: ChartType;
  colorByColumn: string;
  tooltipColumns: string;
  drillDownColumns: string;
  bookmarks: IBookmark[];
  onChange: (partial: {
    colorByColumn?: string;
    tooltipColumns?: string;
    drillDownColumns?: string;
  }) => void;
  onSaveBookmark: (name: string) => void;
  onApplyBookmark: (name: string) => void;
  onDeleteBookmark: (name: string) => void;
}

const DRILL_LEVELS = 3;

const AdvancedOptions: React.FC<IAdvancedOptionsProps> = ({
  columns,
  chartType,
  colorByColumn,
  tooltipColumns,
  drillDownColumns,
  bookmarks,
  onChange,
  onSaveBookmark,
  onApplyBookmark,
  onDeleteBookmark,
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [bookmarkName, setBookmarkName] = React.useState('');
  const idPrefix = React.useRef(`sdv-adv-${Math.random().toString(36).slice(2, 8)}`).current;

  const tooltipList = tooltipColumns
    ? tooltipColumns.split(',').map(s => s.trim()).filter(Boolean)
    : [];
  const drillList = drillDownColumns
    ? drillDownColumns.split(',').map(s => s.trim())
    : [];

  const toggleTooltipColumn = (col: string) => {
    const next = tooltipList.indexOf(col) >= 0
      ? tooltipList.filter(c => c !== col)
      : [...tooltipList, col];
    onChange({ tooltipColumns: next.join(',') });
  };

  const setDrillLevel = (level: number, col: string) => {
    const next = [...drillList];
    while (next.length <= level) next.push('');
    if (!col) {
      // Clearing a level invalidates every level below it too — otherwise
      // e.g. clearing level 1 while level 2 is set leaves a leading empty
      // slot ("",B,C) that the trailing-empty trim below can't remove.
      next.length = level;
    } else {
      next[level] = col;
    }
    // Trim trailing empties so "Level 1 only" produces a clean single-entry list
    while (next.length && !next[next.length - 1]) next.pop();
    onChange({ drillDownColumns: next.join(',') });
  };

  const showColorBy = isScatterOrBubble(chartType);

  return (
    <div className={styles.editPanel}>
      <button
        className={styles.configToggleButton}
        onClick={() => setIsOpen(o => !o)}
        aria-expanded={isOpen}
      >
        {isOpen ? strings.AdvancedOptionsOpen : strings.AdvancedOptionsClosed}
      </button>

      {isOpen && (
        <div>
          {showColorBy && (
            <div className={styles.fieldRow}>
              <div className={styles.fieldGroup}>
                <label htmlFor={`${idPrefix}-colorby`}>{strings.ColorByLabel}</label>
                <select
                  id={`${idPrefix}-colorby`}
                  value={colorByColumn}
                  onChange={e => onChange({ colorByColumn: e.target.value })}
                >
                  <option value="">{strings.NoneOption}</option>
                  {columns.map(col => <option key={col} value={col}>{col}</option>)}
                </select>
                <span className={styles.helpText}>{strings.ColorByHelp}</span>
              </div>
            </div>
          )}

          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup} style={{ flex: '1 1 100%' }}>
              <label>{strings.TooltipColumnsLabel} <span className={styles.helpText}>{strings.TooltipColumnsHelp}</span></label>
              <ul className={styles.multiSelect}>
                {columns.map(col => (
                  <li key={col} className={tooltipList.indexOf(col) >= 0 ? styles.selected : ''}>
                    <label>
                      <input
                        type="checkbox"
                        checked={tooltipList.indexOf(col) >= 0}
                        onChange={() => toggleTooltipColumn(col)}
                      />
                      {col}
                    </label>
                  </li>
                ))}
              </ul>
            </div>
          </div>

          <div className={styles.fieldRow}>
            {Array.from({ length: DRILL_LEVELS }, (_, level) => (
              <div className={styles.fieldGroup} key={level}>
                <label htmlFor={`${idPrefix}-drill-${level}`}>
                  {fmt(strings.DrillLevelLabel, level + 1)}
                </label>
                <select
                  id={`${idPrefix}-drill-${level}`}
                  value={drillList[level] || ''}
                  onChange={e => setDrillLevel(level, e.target.value)}
                  disabled={level > 0 && !drillList[level - 1]}
                >
                  <option value="">{strings.NoneOption}</option>
                  {columns
                    .filter(col => drillList.indexOf(col) < 0 || drillList[level] === col)
                    .map(col => <option key={col} value={col}>{col}</option>)}
                </select>
              </div>
            ))}
          </div>
          <span className={styles.helpText}>{strings.DrillDownHelp}</span>

          <div className={styles.sectionHeader} style={{ marginTop: 12 }}>{strings.BookmarksLabel}</div>
          <div className={styles.fieldRow}>
            <div className={styles.fieldGroup}>
              <label htmlFor={`${idPrefix}-bmname`}>{strings.BookmarkNameLabel}</label>
              <input
                id={`${idPrefix}-bmname`}
                type="text"
                value={bookmarkName}
                onChange={e => setBookmarkName(e.target.value)}
                placeholder={strings.BookmarkNamePlaceholder}
              />
            </div>
            <div className={styles.fieldGroup} style={{ alignSelf: 'flex-end' }}>
              <button
                className={styles.secondaryButton}
                onClick={() => { onSaveBookmark(bookmarkName); setBookmarkName(''); }}
                disabled={!bookmarkName.trim()}
              >
                {strings.SaveBookmarkButton}
              </button>
            </div>
          </div>
          {bookmarks.length > 0 && (
            <ul className={styles.bookmarkList}>
              {bookmarks.map(b => (
                <li key={b.name}>
                  <span className={styles.bookmarkName}>{b.name}</span>
                  <button className={styles.secondaryButton} onClick={() => onApplyBookmark(b.name)}>
                    {strings.ApplyBookmarkButton}
                  </button>
                  <button className={styles.secondaryButton} onClick={() => onDeleteBookmark(b.name)}>
                    {strings.DeleteBookmarkButton}
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
    </div>
  );
};

export default AdvancedOptions;
