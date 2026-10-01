import { WebPartContext } from '@microsoft/sp-webpart-base';
import type { DynamicProperty } from '@microsoft/sp-component-base';
import {
  ChartType,
  DataSourceType,
  SortDirection,
  AggregationType,
  XAxisType,
  TrendlineType,
  ReferenceLineType,
  LegendPosition,
  ThresholdDirection,
} from '../types';

export interface IChartSelection {
  category: string;
  value: number | null;
  series: string;
}

export interface ISharePointSmartChartsWebPartProps {
  // Web part header (above the chart container)
  webPartHeader: string;
  showWebPartHeader: boolean;
  // Core chart settings
  chartType: ChartType;
  chartTitle: string;
  showLegend: boolean;
  showDataTable: boolean;
  stacked: boolean;
  xAxisLabel: string;
  yAxisLabel: string;
  // Layout
  legendPosition: LegendPosition;
  chartHeight: number;
  showExportBar: boolean;
  // Data source
  dataSourceType: DataSourceType;
  uploadedData: string;
  uploadedFileName: string;
  siteUrl: string;
  listName: string;
  dataUrl: string;
  dataPath: string;
  delimiter: string;
  // Column mapping
  xColumn: string;
  yColumns: string;
  labelColumn: string;
  sizeColumn: string;
  // Colors
  colorPalette: string;
  seriesColors: string;
  // Data labels & formatting
  showDataLabels: boolean;
  valuePrefix: string;
  valueSuffix: string;
  valueDecimals: number;
  abbreviateNumbers: boolean;
  // Axes & grid
  yAxisMin: string;
  yAxisMax: string;
  logScale: boolean;
  showGridLines: boolean;
  xLabelRotation: number;
  // Data manipulation (inline controls)
  sortColumn: string;
  sortDirection: SortDirection;
  rowLimit: number;
  filterColumn: string;
  filterValue: string;
  filterOperator: string;
  // Aggregation (inline controls)
  groupByColumn: string;
  aggregation: AggregationType;
  // Data & refresh
  refreshIntervalMinutes: number;
  cacheMinutes: number;
  sheetName: string;
  // Axes
  xAxisType: XAxisType;
  // Combo charts
  seriesTypes: string;
  // Conditional formatting
  thresholdValue: string;
  thresholdDirection: ThresholdDirection;
  thresholdColor: string;
  // Analytics
  trendline: TrendlineType;
  trendWindow: number;
  forecastPeriods: number;
  // Reference line
  referenceLineType: ReferenceLineType;
  referenceLineValue: string;
  referenceLineColor: string;
  // Histogram
  histogramBins: number;
  // Interactivity
  showViewerFilters: boolean;
  detailsOnDemand: boolean;
  drillDownColumns: string;
  // Spotfire-style extras (inline Advanced Options)
  colorByColumn: string;
  tooltipColumns: string;
  bookmarks: string;
  // Axes (advanced)
  logScaleX: boolean;
  logScaleY2: boolean;
  stepLine: boolean;
  // Dual Y axis
  y2Columns: string;
  y2AxisLabel: string;
  // Error bars
  errorBarType: string;
  errorBarColumn: string;
  // Data point overlay on bar charts
  showDataPoints: boolean;
  // Significance annotation brackets
  significancePairs: string;
  // Bubble chart size legend
  showBubbleSizeLegend: boolean;
  // Waterfall
  waterfallShowTotal: boolean;
  waterfallPositiveColor: string;
  waterfallNegativeColor: string;
  waterfallTotalColor: string;
  // Dual axis number formatting
  y2ValuePrefix: string;
  y2ValueSuffix: string;
  // Annotations: "xValue, note" per line
  annotations: string;
  // Filter driven by another web part (Dynamic Data consumer)
  externalFilter?: DynamicProperty<string>;
  externalFilterColumn: string;
}

export interface ISharePointSmartChartsProps extends ISharePointSmartChartsWebPartProps {
  context: WebPartContext;
  /** Resolved value of the Dynamic Data filter source ('' when unset/empty) */
  externalFilterValue: string;
  isDarkTheme: boolean;
  isReadOnly: boolean;
  onPropertiesUpdate: (props: Partial<ISharePointSmartChartsWebPartProps>) => void;
  onItemSelected: (selection: IChartSelection) => void;
}
