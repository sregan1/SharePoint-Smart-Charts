import * as React from 'react';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import { IChartRecord, fmt } from '../types';
import styles from './SharePointSmartCharts.module.scss';

interface IDataTableProps {
  data: IChartRecord[];
  columns: string[];
}

const PAGE_SIZE = 20;

const DataTable: React.FC<IDataTableProps> = ({ data, columns }) => {
  const [page, setPage] = React.useState(0);

  if (!data.length || !columns.length) return null;

  const pageCount = Math.ceil(data.length / PAGE_SIZE);
  // Clamp in case the data shrank (e.g. a filter was applied) while on a later page
  const safePage = Math.min(page, pageCount - 1);
  const pageData = data.slice(safePage * PAGE_SIZE, (safePage + 1) * PAGE_SIZE);

  return (
    <div style={{ marginTop: 16 }}>
      <div className={styles.sectionHeader}>{fmt(strings.DataTableHeader, data.length)}</div>
      <div className={styles.tableWrapper}>
        <table className={styles.dataTable}>
          <thead>
            <tr>
              {columns.map(col => <th key={col}>{col}</th>)}
            </tr>
          </thead>
          <tbody>
            {pageData.map((row, i) => (
              <tr key={safePage * PAGE_SIZE + i}>
                {columns.map(col => (
                  <td key={col} title={String(row[col] ?? '')}>
                    {String(row[col] ?? '')}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {pageCount > 1 && (
        <div className={styles.paginationBar}>
          <button
            className={styles.secondaryButton}
            onClick={() => setPage(Math.max(0, safePage - 1))}
            disabled={safePage === 0}
            style={{ padding: '4px 10px' }}
          >
            {strings.PrevPageButton}
          </button>
          <span>{fmt(strings.PageOfLabel, safePage + 1, pageCount)}</span>
          <button
            className={styles.secondaryButton}
            onClick={() => setPage(Math.min(pageCount - 1, safePage + 1))}
            disabled={safePage === pageCount - 1}
            style={{ padding: '4px 10px' }}
          >
            {strings.NextPageButton}
          </button>
        </div>
      )}
    </div>
  );
};

export default DataTable;
