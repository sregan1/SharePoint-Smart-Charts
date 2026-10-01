import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneLabel,
  PropertyPaneDropdownOptionType,
  PropertyPaneDynamicField,
  PropertyPaneDynamicFieldSet,
  DynamicDataSharedDepth,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import SharePointSmartCharts from './components/SharePointSmartCharts';
import {
  ISharePointSmartChartsProps,
  ISharePointSmartChartsWebPartProps,
  IChartSelection,
} from './components/ISharePointSmartChartsProps';
import * as strings from 'SharePointSmartChartsWebPartStrings';
import { ChartType, supportsDualAxis, supportsErrorBars } from './types';

export default class SharePointSmartChartsWebPart
  extends BaseClientSideWebPart<ISharePointSmartChartsWebPartProps>
  implements IDynamicDataCallables {

  private _isDarkTheme: boolean = false;
  private _partReady: boolean = false;
  private _selection: IChartSelection = { category: '', value: null, series: '' };

  // ---- Dynamic Data source: lets other web parts react to chart clicks ----

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      { id: 'selectedCategory', title: strings.DynamicDataCategoryLabel },
      { id: 'selectedValue', title: strings.DynamicDataValueLabel },
      { id: 'selectedSeries', title: strings.DynamicDataSeriesLabel },
    ];
  }

  public getPropertyValue(propertyId: string): string | number | null {
    switch (propertyId) {
      case 'selectedCategory': return this._selection.category;
      case 'selectedValue': return this._selection.value;
      case 'selectedSeries': return this._selection.series;
    }
    throw new Error(`Unknown dynamic data property: ${propertyId}`);
  }

  private _handleItemSelected = (selection: IChartSelection): void => {
    this._selection = selection;
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedCategory');
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedValue');
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedSeries');
  };

  // Value published by another web part's Dynamic Data source, or '' when unconnected
  private _getExternalFilterValue(): string {
    const dp = this.properties.externalFilter;
    if (!dp) return '';
    const v = dp.tryGetValue();
    return v === undefined || v === null ? '' : String(v);
  }

  public render(): void {
    if (!this._partReady) return; // SPFx renders again once onInit resolves
    const p = this.properties;
    const element: React.ReactElement<ISharePointSmartChartsProps> = React.createElement(
      SharePointSmartCharts,
      {
        // Header
        webPartHeader: p.webPartHeader || '',
        showWebPartHeader: p.showWebPartHeader !== false,
        // Core
        chartType: p.chartType || 'bar',
        chartTitle: p.chartTitle || '',
        showLegend: p.showLegend !== false,
        showDataTable: p.showDataTable || false,
        stacked: p.stacked || false,
        xAxisLabel: p.xAxisLabel || '',
        yAxisLabel: p.yAxisLabel || '',
        // Layout
        legendPosition: p.legendPosition || 'bottom',
        chartHeight: p.chartHeight || 400,
        showExportBar: p.showExportBar !== false,
        // Data source
        dataSourceType: p.dataSourceType || 'upload',
        uploadedData: p.uploadedData || '',
        uploadedFileName: p.uploadedFileName || '',
        siteUrl: p.siteUrl || '',
        listName: p.listName || '',
        dataUrl: p.dataUrl || '',
        dataPath: p.dataPath || '',
        delimiter: p.delimiter || '',
        // Column mapping
        xColumn: p.xColumn || '',
        yColumns: p.yColumns || '',
        labelColumn: p.labelColumn || '',
        sizeColumn: p.sizeColumn || '',
        // Colors
        colorPalette: p.colorPalette || 'office',
        seriesColors: p.seriesColors || '',
        // Data labels & formatting
        showDataLabels: p.showDataLabels || false,
        valuePrefix: p.valuePrefix || '',
        valueSuffix: p.valueSuffix || '',
        valueDecimals: p.valueDecimals !== undefined ? p.valueDecimals : 0,
        abbreviateNumbers: p.abbreviateNumbers || false,
        // Axes & grid
        yAxisMin: p.yAxisMin || '',
        yAxisMax: p.yAxisMax || '',
        logScale: p.logScale || false,
        showGridLines: p.showGridLines !== false,
        xLabelRotation: p.xLabelRotation !== undefined ? p.xLabelRotation : 0,
        // Data manipulation
        sortColumn: p.sortColumn || '',
        sortDirection: p.sortDirection || 'asc',
        rowLimit: p.rowLimit || 0,
        filterColumn: p.filterColumn || '',
        filterValue: p.filterValue || '',
        filterOperator: p.filterOperator || 'contains',
        // Aggregation
        groupByColumn: p.groupByColumn || '',
        aggregation: p.aggregation || 'none',
        // Data & refresh
        refreshIntervalMinutes: p.refreshIntervalMinutes || 0,
        cacheMinutes: p.cacheMinutes || 0,
        sheetName: p.sheetName || '',
        // Axes
        xAxisType: p.xAxisType || 'auto',
        // Combo charts
        seriesTypes: p.seriesTypes || '',
        // Conditional formatting
        thresholdValue: p.thresholdValue || '',
        thresholdDirection: p.thresholdDirection || 'below',
        thresholdColor: p.thresholdColor || '#d13438',
        // Analytics
        trendline: p.trendline || 'none',
        trendWindow: p.trendWindow || 3,
        forecastPeriods: p.forecastPeriods || 0,
        // Reference line
        referenceLineType: p.referenceLineType || 'none',
        referenceLineValue: p.referenceLineValue || '',
        referenceLineColor: p.referenceLineColor || '#666666',
        // Histogram
        histogramBins: p.histogramBins || 10,
        // Interactivity
        showViewerFilters: p.showViewerFilters || false,
        detailsOnDemand: p.detailsOnDemand || false,
        drillDownColumns: p.drillDownColumns || '',
        // Spotfire-style extras
        colorByColumn: p.colorByColumn || '',
        tooltipColumns: p.tooltipColumns || '',
        bookmarks: p.bookmarks || '',
        // Axes (advanced)
        logScaleX: p.logScaleX || false,
        logScaleY2: p.logScaleY2 || false,
        stepLine: p.stepLine || false,
        // Dual Y axis
        y2Columns: p.y2Columns || '',
        y2AxisLabel: p.y2AxisLabel || '',
        // Error bars
        errorBarType: p.errorBarType || 'none',
        errorBarColumn: p.errorBarColumn || '',
        // Data point overlay
        showDataPoints: p.showDataPoints || false,
        // Significance brackets
        significancePairs: p.significancePairs || '',
        // Bubble size legend
        showBubbleSizeLegend: p.showBubbleSizeLegend || false,
        // Waterfall
        waterfallShowTotal: p.waterfallShowTotal || false,
        waterfallPositiveColor: p.waterfallPositiveColor || '',
        waterfallNegativeColor: p.waterfallNegativeColor || '',
        waterfallTotalColor: p.waterfallTotalColor || '',
        // Dual axis number formatting
        y2ValuePrefix: p.y2ValuePrefix || '',
        y2ValueSuffix: p.y2ValueSuffix || '',
        // Annotations
        annotations: p.annotations || '',
        // Dynamic Data consumer
        externalFilterColumn: p.externalFilterColumn || '',
        externalFilterValue: this._getExternalFilterValue(),
        // Framework
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        isReadOnly: this.displayMode === DisplayMode.Read,
        onPropertiesUpdate: (props: Partial<ISharePointSmartChartsWebPartProps>) => {
          Object.assign(this.properties, props);
          // SPFx only writes property changes into the page's saved draft when
          // they flow through the property pane's dirty-bit pipeline. Settings
          // edited via this in-canvas panel bypass that pipeline entirely, so
          // without this call the edits are silently lost on Publish. There is
          // no supported public API for this — _afterPropertyUpdated is the
          // same internal hook the real property pane uses to mark the page
          // dirty and persist the change; it also triggers a re-render.
          const self = this as unknown as { _afterPropertyUpdated?: (shouldRefresh: boolean) => void };
          if (typeof self._afterPropertyUpdated === 'function') {
            self._afterPropertyUpdated(true);
          } else {
            // Internal hook missing (SPFx upgrade?) — edits won't persist on Publish.
            // eslint-disable-next-line no-console
            console.error('SharePoint Smart Charts: _afterPropertyUpdated is unavailable; inline edits may not be saved.');
            this.render();
          }
        },
        onItemSelected: this._handleItemSelected,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.context.dynamicDataSourceManager.initializeSource(this);
    this._partReady = true;
  }

  // Relative luminance of a #rgb/#rrggbb color (0 = black, 1 = white); -1 if unparseable
  private _luminance(color: string | undefined): number {
    const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec((color || '').trim());
    if (!m) return -1;
    const h = m[1].length === 3 ? m[1].split('').map(c => c + c).join('') : m[1];
    const [r, g, b] = [0, 2, 4].map(i => parseInt(h.substr(i, 2), 16) / 255);
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    // "Strong" section variants may be dark without isInverted being set
    const bgLuminance = this._luminance(semanticColors?.bodyBackground);
    this._isDarkTheme = !!currentTheme.isInverted || (bgLuminance >= 0 && bgLuminance < 0.4);
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
    // SPFx raises the first theme change before onInit, when this.properties
    // and this.displayMode aren't safe to read — SPFx renders after onInit anyway.
    if (this._partReady) this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Optional color fields: empty is valid (use the default), otherwise must be
  // a 3/6/8-digit hex color. An invalid CSS color silently degrades the chart
  // (canvas ignores it) with no visible error, so catch it here instead.
  private _validateOptionalColor(value: string): string {
    if (!value || !value.trim()) return '';
    return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(value.trim())
      ? ''
      : strings.ColorValidationError;
  }

  // Optional numeric fields: empty is valid (auto), anything else must parse as a number
  private _validateOptionalNumber(value: string): string {
    if (!value || !value.trim()) return '';
    return isNaN(Number(value)) ? strings.NumericValidationError : '';
  }

  // Fields whose value drives the visibility or disabled state of other pane
  // fields — the pane must rebuild for those states to update live.
  private static readonly PANE_STRUCTURE_FIELDS = [
    'chartType', 'showLegend', 'trendline', 'referenceLineType',
    'errorBarType',
  ];

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (SharePointSmartChartsWebPart.PANE_STRUCTURE_FIELDS.indexOf(propertyPath) >= 0) {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const currentType = this.properties.chartType || 'bar';

    const chartTypes: IPropertyPaneDropdownOption[] = [
      { key: 'header-standard', text: strings.ChartGroupStandard, type: PropertyPaneDropdownOptionType.Header },
      { key: 'bar',           text: strings.ChartTypeBarLabel },
      { key: 'horizontalBar', text: strings.ChartTypeHorizontalBarLabel },
      { key: 'line',          text: strings.ChartTypeLineLabel },
      { key: 'area',          text: strings.ChartTypeAreaLabel },
      { key: 'pie',           text: strings.ChartTypePieLabel },
      { key: 'doughnut',      text: strings.ChartTypeDoughnutLabel },
      { key: 'scatter',       text: strings.ChartTypeScatterLabel },
      { key: 'bubble',        text: strings.ChartTypeBubbleLabel },
      { key: 'radar',         text: strings.ChartTypeRadarLabel },
      { key: 'divider-1',     text: '-', type: PropertyPaneDropdownOptionType.Divider },
      { key: 'header-special', text: strings.ChartGroupSpecialized, type: PropertyPaneDropdownOptionType.Header },
      { key: 'kpi',           text: strings.ChartTypeKpiLabel },
      { key: 'histogram',     text: strings.ChartTypeHistogramLabel },
      { key: 'waterfall',     text: strings.ChartTypeWaterfallLabel },
      { key: 'boxplot',       text: strings.ChartTypeBoxplotLabel },
      { key: 'violin',        text: strings.ChartTypeViolinLabel },
      { key: 'beforeAfter',   text: strings.ChartTypeBeforeAfterLabel },
      { key: 'treemap',       text: strings.ChartTypeTreemapLabel },
      { key: 'heatmap',       text: strings.ChartTypeHeatmapLabel },
    ];

    const legendPositions = [
      { key: 'top', text: strings.LegendTopLabel },
      { key: 'bottom', text: strings.LegendBottomLabel },
      { key: 'left', text: strings.LegendLeftLabel },
      { key: 'right', text: strings.LegendRightLabel },
    ];

    const colorPalettes = [
      { key: 'office', text: strings.PaletteOfficeLabel },
      { key: 'colorblind', text: strings.PaletteColorblindLabel },
      { key: 'vibrant', text: strings.PaletteVibrantLabel },
      { key: 'pastel', text: strings.PalettePastelLabel },
      { key: 'monochrome', text: strings.PaletteMonochromeLabel },
      { key: 'trafficLight', text: strings.PaletteTrafficLightLabel },
      { key: 'warm', text: strings.PaletteWarmLabel },
      { key: 'cool', text: strings.PaletteCoolLabel },
    ];

    const headerGroup = {
      groupName: strings.HeaderGroupName,
      groupFields: [
        PropertyPaneToggle('showWebPartHeader', {
          label: strings.ShowWebPartHeaderFieldLabel,
          checked: this.properties.showWebPartHeader !== false,
        }),
        PropertyPaneTextField('webPartHeader', {
          label: strings.WebPartHeaderFieldLabel,
          placeholder: strings.WebPartHeaderPlaceholder,
        }),
      ],
    };

    const chartSettingsGroup = {
      groupName: strings.ChartSettingsGroupName,
      groupFields: [
        PropertyPaneTextField('chartTitle', {
          label: strings.ChartTitleFieldLabel,
        }),
        PropertyPaneDropdown('chartType', {
          label: strings.ChartTypeFieldLabel,
          options: chartTypes,
          selectedKey: currentType,
        }),
        PropertyPaneDropdown('legendPosition', {
          label: strings.LegendPositionFieldLabel,
          options: legendPositions,
          selectedKey: this.properties.legendPosition || 'bottom',
          disabled: this.properties.showLegend === false,
        }),
        PropertyPaneSlider('chartHeight', {
          label: strings.ChartHeightFieldLabel,
          min: 150,
          max: 1200,
          step: 50,
          value: this.properties.chartHeight || 400,
        }),
        PropertyPaneToggle('showLegend', {
          label: strings.ShowLegendFieldLabel,
          checked: this.properties.showLegend !== false,
        }),
        ...(['bar', 'horizontalBar', 'line', 'area'].indexOf(currentType) >= 0 ? [
          PropertyPaneToggle('stacked', {
            label: strings.StackedFieldLabel,
            checked: this.properties.stacked || false,
          }),
        ] : []),
        PropertyPaneToggle('showDataTable', {
          label: strings.ShowDataTableFieldLabel,
          checked: this.properties.showDataTable || false,
        }),
        PropertyPaneToggle('showExportBar', {
          label: strings.ShowExportBarFieldLabel,
          checked: this.properties.showExportBar !== false,
        }),
        ...(['pie', 'doughnut', 'kpi', 'treemap', 'heatmap'].indexOf(currentType) < 0 ? [
          PropertyPaneTextField('xAxisLabel', {
            label: strings.XAxisLabelFieldLabel,
            placeholder: strings.XAxisPlaceholder,
          }),
          PropertyPaneTextField('yAxisLabel', {
            label: strings.YAxisLabelFieldLabel,
            placeholder: strings.YAxisPlaceholder,
          }),
        ] : []),
        ...(currentType === 'histogram' ? [
          PropertyPaneSlider('histogramBins', {
            label: strings.HistogramBinsFieldLabel,
            min: 4,
            max: 50,
            step: 1,
            value: this.properties.histogramBins || 10,
          }),
        ] : []),
      ],
    };

    const colorsGroup = {
      groupName: strings.ColorsGroupName,
      groupFields: [
        PropertyPaneDropdown('colorPalette', {
          label: strings.ColorPaletteFieldLabel,
          options: colorPalettes,
          selectedKey: this.properties.colorPalette || 'office',
        }),
      ],
    };

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [headerGroup, chartSettingsGroup],
        },
        {
          header: { description: strings.AppearancePageDescription },
          displayGroupsAsAccordion: true,
          groups: [
            colorsGroup,
            {
              groupName: strings.DataLabelsGroupName,
              groupFields: [
                PropertyPaneToggle('showDataLabels', {
                  label: strings.ShowDataLabelsFieldLabel,
                  checked: this.properties.showDataLabels || false,
                }),
                PropertyPaneTextField('valuePrefix', {
                  label: strings.ValuePrefixFieldLabel,
                  placeholder: strings.ValuePrefixPlaceholder,
                }),
                PropertyPaneTextField('valueSuffix', {
                  label: strings.ValueSuffixFieldLabel,
                  placeholder: strings.ValueSuffixPlaceholder,
                }),
                PropertyPaneSlider('valueDecimals', {
                  label: strings.ValueDecimalsFieldLabel,
                  min: 0,
                  max: 4,
                  step: 1,
                  value: this.properties.valueDecimals !== undefined ? this.properties.valueDecimals : 0,
                }),
                PropertyPaneToggle('abbreviateNumbers', {
                  label: strings.AbbreviateNumbersFieldLabel,
                  checked: this.properties.abbreviateNumbers || false,
                }),
              ],
            },
            ...(['pie', 'doughnut', 'kpi', 'treemap', 'heatmap', 'radar'].indexOf(currentType) < 0 ? [{
              groupName: strings.AxesGridGroupName,
              groupFields: [
                ...(['bar', 'horizontalBar', 'line', 'area', 'waterfall', 'histogram', 'boxplot', 'violin', 'beforeAfter'].indexOf(currentType) >= 0 ? [
                  PropertyPaneTextField('yAxisMin', {
                    label: strings.YAxisMinFieldLabel,
                    placeholder: strings.AutoPlaceholder,
                    onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                  }),
                  PropertyPaneTextField('yAxisMax', {
                    label: strings.YAxisMaxFieldLabel,
                    placeholder: strings.AutoPlaceholder,
                    onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                  }),
                ] : []),
                PropertyPaneToggle('logScale', {
                  label: strings.LogScaleFieldLabel,
                  checked: this.properties.logScale || false,
                }),
                ...(['scatter', 'bubble'].indexOf(currentType) >= 0 ? [
                  PropertyPaneToggle('logScaleX', {
                    label: strings.LogScaleXFieldLabel,
                    checked: this.properties.logScaleX || false,
                  }),
                ] : []),
                ...(['line', 'area'].indexOf(currentType) >= 0 ? [
                  PropertyPaneToggle('stepLine', {
                    label: strings.StepLineFieldLabel,
                    checked: this.properties.stepLine || false,
                  }),
                ] : []),
                PropertyPaneToggle('showGridLines', {
                  label: strings.ShowGridLinesFieldLabel,
                  checked: this.properties.showGridLines !== false,
                }),
                PropertyPaneSlider('xLabelRotation', {
                  label: strings.XLabelRotationFieldLabel,
                  min: 0,
                  max: 90,
                  step: 15,
                  value: this.properties.xLabelRotation !== undefined ? this.properties.xLabelRotation : 0,
                }),
                ...(['bar', 'line', 'area'].indexOf(currentType) >= 0 ? [
                  PropertyPaneDropdown('xAxisType', {
                    label: strings.XAxisTypeFieldLabel,
                    options: [
                      { key: 'auto', text: strings.XAxisTypeAuto },
                      { key: 'category', text: strings.XAxisTypeCategory },
                      { key: 'time', text: strings.XAxisTypeTime },
                    ],
                    selectedKey: this.properties.xAxisType || 'auto',
                  }),
                ] : []),
              ],
            }] : []),
          ],
        },
        {
          header: { description: strings.AdvancedPageDescription },
          displayGroupsAsAccordion: true,
          groups: [
            ...(['bar', 'horizontalBar', 'line', 'area', 'scatter'].indexOf(currentType) >= 0 ? [
            {
              groupName: strings.AnalyticsGroupName,
              groupFields: [
                PropertyPaneDropdown('trendline', {
                  label: strings.TrendlineFieldLabel,
                  options: [
                    { key: 'none', text: strings.TrendlineNone },
                    { key: 'linear', text: strings.TrendlineLinear },
                    { key: 'movingAverage', text: strings.TrendlineMovingAverage },
                  ],
                  selectedKey: this.properties.trendline || 'none',
                }),
                PropertyPaneSlider('trendWindow', {
                  label: strings.TrendWindowFieldLabel,
                  min: 2,
                  max: 20,
                  step: 1,
                  value: this.properties.trendWindow || 3,
                  disabled: this.properties.trendline !== 'movingAverage',
                }),
                PropertyPaneSlider('forecastPeriods', {
                  label: strings.ForecastPeriodsFieldLabel,
                  min: 0,
                  max: 12,
                  step: 1,
                  value: this.properties.forecastPeriods || 0,
                  disabled: this.properties.trendline !== 'linear',
                }),
              ],
            },
            ] : []),
            ...(['bar', 'horizontalBar', 'line', 'area'].indexOf(currentType) >= 0 ? [
            {
              groupName: strings.ReferenceLineGroupName,
              groupFields: [
                PropertyPaneDropdown('referenceLineType', {
                  label: strings.ReferenceLineTypeFieldLabel,
                  options: [
                    { key: 'none', text: strings.TrendlineNone },
                    { key: 'fixed', text: strings.RefLineFixed },
                    { key: 'mean', text: strings.RefLineMean },
                    { key: 'median', text: strings.RefLineMedian },
                  ],
                  selectedKey: this.properties.referenceLineType || 'none',
                }),
                PropertyPaneTextField('referenceLineValue', {
                  label: strings.ReferenceLineValueFieldLabel,
                  placeholder: strings.AutoPlaceholder,
                  disabled: this.properties.referenceLineType !== 'fixed',
                  onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                }),
                PropertyPaneTextField('referenceLineColor', {
                  label: strings.ReferenceLineColorFieldLabel,
                  placeholder: '#666666',
                  onGetErrorMessage: (value: string) => this._validateOptionalColor(value),
                }),
              ],
            },
            ] : []),
            ...(supportsDualAxis(currentType as ChartType) ? [
            {
              groupName: strings.DualAxisGroupName,
              groupFields: [
                PropertyPaneTextField('y2Columns', {
                  label: strings.Y2ColumnsFieldLabel,
                  placeholder: strings.Y2ColumnsPlaceholder,
                  value: this.properties.y2Columns || '',
                }),
                PropertyPaneTextField('y2AxisLabel', {
                  label: strings.Y2AxisLabelFieldLabel,
                  value: this.properties.y2AxisLabel || '',
                }),
                PropertyPaneToggle('logScaleY2', {
                  label: strings.LogScaleY2FieldLabel,
                  checked: this.properties.logScaleY2 || false,
                }),
                PropertyPaneTextField('y2ValuePrefix', {
                  label: strings.Y2ValuePrefixFieldLabel,
                  placeholder: strings.ValuePrefixPlaceholder,
                }),
                PropertyPaneTextField('y2ValueSuffix', {
                  label: strings.Y2ValueSuffixFieldLabel,
                  placeholder: strings.ValueSuffixPlaceholder,
                }),
              ],
            },
            ] : []),
            ...(supportsErrorBars(currentType as ChartType) ? [
            {
              groupName: strings.ErrorBarsGroupName,
              groupFields: [
                PropertyPaneDropdown('errorBarType', {
                  label: strings.ErrorBarTypeFieldLabel,
                  options: [
                    { key: 'none', text: strings.ErrorBarTypeNone },
                    { key: 'custom', text: strings.ErrorBarTypeCustom },
                    { key: 'sd', text: strings.ErrorBarTypeSd },
                    { key: 'sem', text: strings.ErrorBarTypeSem },
                  ],
                  selectedKey: this.properties.errorBarType || 'none',
                }),
                PropertyPaneTextField('errorBarColumn', {
                  label: strings.ErrorBarColumnFieldLabel,
                  value: this.properties.errorBarColumn || '',
                  disabled: (this.properties.errorBarType || 'none') !== 'custom',
                }),
              ],
            },
            ] : []),
            ...(currentType === 'waterfall' ? [
            {
              groupName: strings.WaterfallGroupName,
              groupFields: [
                PropertyPaneToggle('waterfallShowTotal', {
                  label: strings.WaterfallShowTotalFieldLabel,
                  checked: this.properties.waterfallShowTotal || false,
                }),
                PropertyPaneTextField('waterfallPositiveColor', {
                  label: strings.WaterfallPositiveColorFieldLabel,
                  placeholder: strings.AutoPlaceholder,
                  onGetErrorMessage: (value: string) => this._validateOptionalColor(value),
                }),
                PropertyPaneTextField('waterfallNegativeColor', {
                  label: strings.WaterfallNegativeColorFieldLabel,
                  placeholder: '#d13438',
                  onGetErrorMessage: (value: string) => this._validateOptionalColor(value),
                }),
                PropertyPaneTextField('waterfallTotalColor', {
                  label: strings.WaterfallTotalColorFieldLabel,
                  placeholder: '#69797e',
                  onGetErrorMessage: (value: string) => this._validateOptionalColor(value),
                }),
              ],
            },
            ] : []),
            ...(['bar', 'line', 'area'].indexOf(currentType) >= 0 ? [
            {
              groupName: strings.AnnotationsGroupName,
              groupFields: [
                PropertyPaneTextField('annotations', {
                  label: strings.AnnotationsFieldLabel,
                  placeholder: strings.AnnotationsPlaceholder,
                  description: strings.AnnotationsHelp,
                  multiline: true,
                  rows: 4,
                }),
              ],
            },
            ] : []),
            {
              groupName: strings.ExternalFilterGroupName,
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: strings.ExternalFilterFieldLabel,
                  fields: [PropertyPaneDynamicField('externalFilter', { label: strings.ExternalFilterFieldLabel })],
                  sharedConfiguration: { depth: DynamicDataSharedDepth.Property },
                }),
                PropertyPaneTextField('externalFilterColumn', {
                  label: strings.ExternalFilterColumnFieldLabel,
                  placeholder: strings.ExternalFilterColumnPlaceholder,
                }),
              ],
            },
            ...(currentType === 'bar' ? [
            {
              groupName: strings.SignificanceGroupName,
              groupFields: [
                PropertyPaneTextField('significancePairs', {
                  label: strings.SignificancePairsFieldLabel,
                  placeholder: strings.SignificancePairsPlaceholder,
                  description: strings.SignificancePairsHelp,
                  multiline: true,
                  rows: 3,
                  value: this.properties.significancePairs || '',
                }),
              ],
            },
            ] : []),
            {
              groupName: strings.InteractivityGroupName,
              groupFields: [
                PropertyPaneToggle('showViewerFilters', {
                  label: strings.ShowViewerFiltersFieldLabel,
                  checked: this.properties.showViewerFilters || false,
                }),
                PropertyPaneToggle('detailsOnDemand', {
                  label: strings.DetailsOnDemandFieldLabel,
                  checked: this.properties.detailsOnDemand || false,
                }),
                ...(['bar', 'horizontalBar'].indexOf(currentType) >= 0 ? [
                  PropertyPaneToggle('showDataPoints', {
                    label: strings.ShowDataPointsFieldLabel,
                    checked: this.properties.showDataPoints || false,
                  }),
                ] : []),
                ...(currentType === 'bubble' ? [
                  PropertyPaneToggle('showBubbleSizeLegend', {
                    label: strings.ShowBubbleSizeLegendFieldLabel,
                    checked: this.properties.showBubbleSizeLegend || false,
                  }),
                ] : []),
              ],
            },
            ...(['bar', 'horizontalBar', 'line', 'area', 'scatter', 'bubble', 'histogram', 'kpi'].indexOf(currentType) >= 0 ? [
            {
              groupName: strings.ConditionalGroupName,
              groupFields: [
                PropertyPaneTextField('thresholdValue', {
                  label: strings.ThresholdValueFieldLabel,
                  placeholder: strings.ThresholdValuePlaceholder,
                  onGetErrorMessage: (value: string) => this._validateOptionalNumber(value),
                }),
                PropertyPaneDropdown('thresholdDirection', {
                  label: strings.ThresholdDirectionFieldLabel,
                  options: [
                    { key: 'below', text: strings.ThresholdBelow },
                    { key: 'above', text: strings.ThresholdAbove },
                  ],
                  selectedKey: this.properties.thresholdDirection || 'below',
                }),
                PropertyPaneTextField('thresholdColor', {
                  label: strings.ThresholdColorFieldLabel,
                  placeholder: '#d13438',
                  onGetErrorMessage: (value: string) => this._validateOptionalColor(value),
                }),
              ],
            },
            ] : []),
            {
              groupName: strings.DataRefreshGroupName,
              groupFields: [
                PropertyPaneSlider('refreshIntervalMinutes', {
                  label: strings.RefreshIntervalFieldLabel,
                  min: 0,
                  max: 60,
                  step: 5,
                  value: this.properties.refreshIntervalMinutes || 0,
                  disabled: (this.properties.dataSourceType || 'upload') === 'upload',
                }),
                PropertyPaneSlider('cacheMinutes', {
                  label: strings.CacheMinutesFieldLabel,
                  min: 0,
                  max: 60,
                  step: 5,
                  value: this.properties.cacheMinutes || 0,
                  // Every network loader consults the session cache; uploads don't.
                  disabled: (this.properties.dataSourceType || 'upload') === 'upload',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
