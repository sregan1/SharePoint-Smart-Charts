import * as React from 'react';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import styles from './SharePointSmartCharts.module.scss';

interface IExportBarProps {
  onExportPng: () => void;
  onExportJpeg: () => void;
  onExportCsv: () => void;
  onExportExcel: () => void;
  hasData: boolean;
  hasChart: boolean;
}

const ExportBar: React.FC<IExportBarProps> = ({
  onExportPng,
  onExportJpeg,
  onExportCsv,
  onExportExcel,
  hasData,
  hasChart,
}) => (
  <div className={styles.exportBar}>
    <span className={styles.exportLabel}>{strings.ExportLabel}</span>
    <button
      className={styles.exportButton}
      onClick={onExportPng}
      disabled={!hasChart}
      title={strings.ExportPngTitle}
    >
      <span className={styles.exportIcon} aria-hidden="true">🖼</span> PNG
    </button>
    <button
      className={styles.exportButton}
      onClick={onExportJpeg}
      disabled={!hasChart}
      title={strings.ExportJpegTitle}
    >
      <span className={styles.exportIcon} aria-hidden="true">📷</span> JPEG
    </button>
    <button
      className={styles.exportButton}
      onClick={onExportCsv}
      disabled={!hasData}
      title={strings.ExportCsvTitle}
    >
      <span className={styles.exportIcon} aria-hidden="true">📄</span> CSV
    </button>
    <button
      className={styles.exportButton}
      onClick={onExportExcel}
      disabled={!hasData}
      title={strings.ExportExcelTitle}
    >
      <span className={styles.exportIcon} aria-hidden="true">📊</span> Excel
    </button>
  </div>
);

export default ExportBar;
