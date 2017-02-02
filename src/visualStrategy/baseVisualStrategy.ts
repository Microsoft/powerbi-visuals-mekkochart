/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual.visualStrategy {
    // d3
    import Axis = d3.svg.Axis;
    import Selection = d3.Selection;
    import LinearScale = d3.scale.Linear;
    import UpdateSelection = d3.selection.Update;

    // powerbi.visuals
    import ISelectionId = powerbi.visuals.ISelectionId;

    // powerbi.extensibility.utils.svg
    import IRect = powerbi.extensibility.utils.svg.IRect;
    import IMargin = powerbi.extensibility.utils.svg.IMargin;
    import ClassAndSelector = powerbi.extensibility.utils.svg.CssConstants.ClassAndSelector;
    import createClassAndSelector = powerbi.extensibility.utils.svg.CssConstants.createClassAndSelector;

    // powerbi.extensibility.utils.chart
    import AxisHelper = powerbi.extensibility.utils.chart.axis;
    import IAxisProperties = AxisHelper.IAxisProperties;
    import CreateAxisOptionsBase = AxisHelper.CreateAxisOptions;
    import getLabelPrecision = powerbi.extensibility.utils.chart.dataLabel.utils.getLabelPrecision;
    import hundredPercentFormat = powerbi.extensibility.utils.chart.dataLabel.utils.hundredPercentFormat;
    import VisualDataLabelsSettings = powerbi.extensibility.utils.chart.dataLabel.VisualDataLabelsSettings;
    import IColumnFormatterCacheManager = powerbi.extensibility.utils.chart.dataLabel.IColumnFormatterCacheManager;
    import createColumnFormatterCacheManager = powerbi.extensibility.utils.chart.dataLabel.utils.createColumnFormatterCacheManager;

    // powerbi.extensibility.utils.interactivity
    import IInteractivityService = powerbi.extensibility.utils.interactivity.IInteractivityService;

    // powerbi.extensibility.utils.formatting
    import TextProperties = powerbi.extensibility.utils.formatting.TextProperties;
    import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;
    import IValueFormatter = powerbi.extensibility.utils.formatting.IValueFormatter;
    import textMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;

    // powerbi.extensibility.utils.type
    import ValueType = powerbi.extensibility.utils.type.ValueType;
    import PixelConverter = powerbi.extensibility.utils.type.PixelConverter;

    interface LayoutFunction {
        (dataPoint: MekkoChartColumnDataPoint): number;
    }

    export class BaseVisualStrategy implements IVisualStrategy {
        private static ItemSelector: ClassAndSelector = createClassAndSelector("column");
        private static BorderSelector: ClassAndSelector = createClassAndSelector("mekkoborder");
        private static InteractiveHoverLineSelector: ClassAndSelector = createClassAndSelector("interactive-hover-line");
        private static DragHandleSelector: ClassAndSelector = createClassAndSelector("drag-handle");

        private static TickLabelPaddingFactor: number = 2;
        private static ColumnDataPointValueDelimiter: number = 2;

        private static DefaultInnerPaddingRatio: number = 1;

        private static CircleRadius: number = 6;

        private static CategoryWidthDelimiter: number = 2;

        private static DefaultLabelFillColor: string = "#ffffff";

        private static PercentageFormat: string = "#,0.##%";

        private layout: IMekkoColumnLayout;
        private data: MekkoColumnChartData;
        private graphicsContext: MekkoColumnChartContext;
        private width: number;
        private height: number;
        private margin: IMargin;
        private xProps: IAxisProperties;
        private yProps: IAxisProperties;
        private categoryLayout: MekkoChartCategoryLayout;
        private columnsCenters: number[];
        private columnSelectionLineHandle: Selection<any>;

        private interactivityService: IInteractivityService;
        private viewportHeight: number;

        private static validLabelPositions = [1];

        public setupVisualProps(columnChartProps: MekkoColumnChartContext): void {
            this.graphicsContext = columnChartProps;
            this.margin = columnChartProps.margin;
            this.width = this.graphicsContext.width;
            this.height = this.graphicsContext.height;
            this.categoryLayout = columnChartProps.layout;

            this.interactivityService = columnChartProps.interactivityService;
            this.viewportHeight = columnChartProps.viewportHeight;
        }

        public setData(data: MekkoColumnChartData): void {
            this.data = data;
        }

        private static createFormatter(
            scaleDomain: any[],
            dataDomain: any[],
            dataType,
            isScalar: boolean,
            formatString: string,
            bestTickCount: number,
            tickValues: any[],
            getValueFn: any,
            useTickIntervalForDisplayUnits: boolean = false): IValueFormatter {

            let formatter: IValueFormatter;

            if (dataType.dateTime) {
                if (isScalar) {
                    let value: Date = new Date(scaleDomain[0]),
                        value2: Date = new Date(scaleDomain[1]);

                    if (bestTickCount === 1) {
                        value = value2 = new Date(dataDomain[0]);
                    }

                    formatter = valueFormatter.create({
                        format: formatString,
                        value: value,
                        value2: value2,
                        tickCount: bestTickCount
                    });
                }
                else {
                    const minDate: Date = getValueFn(0, dataType),
                        maxDate: Date = getValueFn(scaleDomain.length - 1, dataType);

                    formatter = valueFormatter.create({
                        format: formatString,
                        value: minDate,
                        value2: maxDate,
                        tickCount: bestTickCount
                    });
                }
            }
            else {
                if (useTickIntervalForDisplayUnits && isScalar && tickValues.length > 1) {
                    const domainMin: number = tickValues[1] - tickValues[0];

                    formatter = valueFormatter.create({
                        format: formatString,
                        value: domainMin,
                        value2: 0,
                        allowFormatBeautification: true
                    });
                }
                else {
                    formatter = valueFormatter.createDefaultFormatter(formatString, true);
                }
            }

            return formatter;
        }

        /**
         * Format the linear tick labels or the category labels.
         */
        private static formatAxisTickValues(
            axis: Axis,
            tickValues: any[],
            formatter: IValueFormatter,
            dataType: ValueType,
            isScalar: boolean,
            getValueFn?: (index: number, type: ValueType) => any) {

            let formattedTickValues: any[] = [];

            if (formatter) {
                if (getValueFn && !(dataType.numeric && isScalar)) {
                    axis.tickFormat((tickValue: any) => {
                        return formatter.format(getValueFn(tickValue, dataType));
                    });

                    formattedTickValues = tickValues.map((tickValue: any) => {
                        return formatter.format(getValueFn(tickValue, dataType));
                    });
                }
                else {
                    axis.tickFormat((tickValue: any) => {
                        return formatter.format(tickValue);
                    });

                    formattedTickValues = tickValues.map((tickValue: any) => {
                        return formatter.format(tickValue);
                    });
                }
            }
            else {
                formattedTickValues = tickValues.map((tickValue: any) => {
                    return getValueFn(tickValue, dataType);
                });
            }

            return formattedTickValues;
        }

        /**
         * Create a D3 axis including scale. Can be vertical or horizontal, and either datetime, numeric, or text.
         * @param options The properties used to create the axis.
         */
        private createAxis(options: CreateAxisOptions): IAxisProperties {
            const pixelSpan: number = options.pixelSpan,
                dataDomain: number[] = options.dataDomain,
                metaDataColumn: DataViewMetadataColumn = options.metaDataColumn,
                outerPadding: number = options.outerPadding || 0,
                isCategoryAxis: boolean = !!options.isCategoryAxis,
                isScalar: boolean = !!options.isScalar,
                isVertical: boolean = !!options.isVertical,
                useTickIntervalForDisplayUnits: boolean = !!options.useTickIntervalForDisplayUnits,
                getValueFn: (index: number, type: ValueType) => any = options.getValueFn,
                categoryThickness: number = options.categoryThickness,
                formatString: string = valueFormatter.getFormatStringByColumn(metaDataColumn),
                dataType: ValueType = AxisHelper.getCategoryValueType(metaDataColumn, isScalar),
                isLogScaleAllowed: boolean = AxisHelper.isLogScalePossible(dataDomain, dataType),
                scale: LinearScale<number, number> = d3.scale.linear(),
                scaleDomain: number[] = [0, 1],
                bestTickCount: number = dataDomain.length || 1,
                borderWidth: number = columnChart.BaseColumnChart.getBorderWidth(options.borderSettings);

            let chartWidth: number = pixelSpan - borderWidth * (bestTickCount - 1);

            if (chartWidth < MekkoChart.MinOrdinalRectThickness) {
                chartWidth = MekkoChart.MinOrdinalRectThickness;
            }

            scale
                .domain(scaleDomain)
                .range([0, chartWidth]);

            const formatter: IValueFormatter = BaseVisualStrategy.createFormatter(
                scaleDomain,
                dataDomain,
                dataType,
                isScalar,
                formatString,
                bestTickCount,
                dataDomain,
                getValueFn,
                useTickIntervalForDisplayUnits);

            const axis: Axis = d3.svg.axis()
                .scale(scale)
                .tickSize(6, 0)
                .orient(isVertical
                    ? "left"
                    : "bottom")
                .ticks(bestTickCount)
                .tickValues(dataDomain);

            let formattedTickValues: any[] = [];

            if (metaDataColumn) {
                formattedTickValues = BaseVisualStrategy.formatAxisTickValues(
                    axis,
                    dataDomain,
                    formatter,
                    dataType,
                    isScalar,
                    getValueFn);
            }

            let xLabelMaxWidth: any;

            if (!isScalar && categoryThickness) {
                xLabelMaxWidth = Math.max(
                    1,
                    categoryThickness - MekkoChart.TickLabelPadding * BaseVisualStrategy.TickLabelPaddingFactor);
            }
            else {
                const labelAreaCount: number = dataDomain.length > 1
                    ? dataDomain.length + 1
                    : dataDomain.length;

                xLabelMaxWidth = labelAreaCount > 1
                    ? pixelSpan / labelAreaCount
                    : pixelSpan;

                xLabelMaxWidth = Math.max(
                    1,
                    xLabelMaxWidth - MekkoChart.TickLabelPadding * BaseVisualStrategy.TickLabelPaddingFactor);
            }

            return {
                scale,
                axis,
                formatter,
                isCategoryAxis,
                xLabelMaxWidth,
                categoryThickness,
                outerPadding,
                isLogScaleAllowed,
                values: formattedTickValues,
                axisType: dataType,
                axisLabel: null,
                usingDefaultDomain: false
            };
        }

        private getCategoryAxis(
            data: MekkoColumnChartData,
            size: number,
            layout: MekkoChartCategoryLayout,
            isVertical: boolean,
            forcedXMin?: DataViewPropertyValue,
            forcedXMax?: DataViewPropertyValue,
            axisScaleType?: string): IAxisProperties {

            const categoryThickness: number = layout.categoryThickness,
                isScalar: boolean = layout.isScalar,
                outerPaddingRatio: number = layout.outerPaddingRatio,
                dataWrapper: DataWrapper = new DataWrapper(data, isScalar);

            let domain: number[] = [];

            if (data.series
                && (data.series.length > 0)
                && data.series[0].data
                && (data.series[0].data.length > 0)) {

                const domainDoubles: number[] = data.series[0].data
                    .map((item: MekkoChartColumnDataPoint) => {
                        return item.originalPosition + item.value / BaseVisualStrategy.ColumnDataPointValueDelimiter;
                    });

                domain = domainDoubles.filter((item: number, position: number) => {
                    return domainDoubles.indexOf(item) === position;
                });
            }

            const axisProperties: IAxisProperties = this.createAxis({
                isScalar,
                isVertical,
                formatString: undefined,
                pixelSpan: size,
                dataDomain: domain,
                metaDataColumn: data.categoryMetadata,
                outerPadding: categoryThickness * outerPaddingRatio,
                isCategoryAxis: true,
                categoryThickness: categoryThickness,
                useTickIntervalForDisplayUnits: true,
                getValueFn: (index: number, type: ValueType) => {
                    const domainIndex: number = domain.indexOf(index),
                        value: number = dataWrapper.lookupXValue(domainIndex, type);

                    return value;
                },
                scaleType: axisScaleType,
                borderSettings: data.borderSettings
            });

            // intentionally updating the input layout by ref
            layout.categoryThickness = axisProperties.categoryThickness;

            return axisProperties;
        }

        public setXScale(
            is100Pct: boolean,
            forcedTickCount?: number,
            forcedXDomain?: any[],
            axisScaleType?: string): IAxisProperties {

            let forcedXMin: any,
                forcedXMax: any;

            if (forcedXDomain && forcedXDomain.length === 2) {
                forcedXMin = forcedXDomain[0];
                forcedXMax = forcedXDomain[1];
            }

            const properties: IAxisProperties = this.xProps = this.getCategoryAxis(
                this.data,
                this.width,
                this.categoryLayout,
                false,
                forcedXMin,
                forcedXMax,
                axisScaleType);

            return properties;
        }

        public setYScale(
            is100Pct: boolean,
            forcedTickCount?: number,
            forcedYDomain?: any[],
            axisScaleType?: string): IAxisProperties {

            const height: number = this.viewportHeight,
                valueDomain: ValueRange<number> = utils.calcValueDomain(this.data.series, is100Pct),
                valueDomainArr: number[] = [valueDomain.min, valueDomain.max],
                combinedDomain: any[] = AxisHelper.combineDomain(forcedYDomain, valueDomainArr),
                shouldClamp: boolean = AxisHelper.scaleShouldClamp(combinedDomain, valueDomainArr),
                metadataColumn: DataViewMetadataColumn = this.data.valuesMetadata[0];

            const formatString: string = is100Pct
                ? BaseVisualStrategy.PercentageFormat
                : valueFormatter.getFormatStringByColumn(metadataColumn);

            const createAxisOptions: MekkoCreateAxisOptions = {
                pixelSpan: height,
                dataDomain: combinedDomain,
                metaDataColumn: metadataColumn,
                formatString: formatString,
                outerPadding: 0,
                isScalar: true,
                isVertical: true,
                forcedTickCount: forcedTickCount,
                useTickIntervalForDisplayUnits: true,
                isCategoryAxis: false,
                scaleType: axisScaleType,
                axisDisplayUnits: 0,
                axisPrecision: 0,
                is100Pct: is100Pct,
                shouldClamp: shouldClamp,
                formatStringProp: undefined,
            };

            this.yProps = AxisHelper.createAxis(createAxisOptions);

            return this.yProps;
        }

        public drawColumns(useAnimation: boolean): MekkoChartDrawInfo {
            const data: MekkoColumnChartData = this.data;

            this.columnsCenters = null;

            const axisOptions: MekkoColumnAxisOptions = {
                columnWidth: 0,
                xScale: this.xProps.scale,
                yScale: this.yProps.scale,
                isScalar: this.categoryLayout.isScalar,
                margin: this.margin,
            };

            const stackedColumnLayout: IMekkoColumnLayout = BaseVisualStrategy.getLayout(data, axisOptions);

            this.layout = stackedColumnLayout;

            const labelDataPoints: LabelDataPoint[] = this.createMekkoLabelDataPoints(),
                series: UpdateSelection<MekkoChartSeries> = utils.drawSeries(
                    data,
                    this.graphicsContext.mainGraphicsContext,
                    axisOptions);

            let shapes: UpdateSelection<MekkoChartColumnDataPoint>;

            if (!useAnimation) {
                shapes = BaseVisualStrategy.drawDefaultShapes(data,
                    series,
                    stackedColumnLayout,
                    BaseVisualStrategy.ItemSelector,
                    this.interactivityService && this.interactivityService.hasSelection());
            }

            utils.applyInteractivity(shapes, this.graphicsContext.onDragStart);

            return {
                axisOptions,
                labelDataPoints,
                shapesSelection: shapes,
                viewport: {
                    height: this.height,
                    width: this.width
                }
            };
        }

        private static drawDefaultShapes(
            data: MekkoColumnChartData,
            series: UpdateSelection<any>,
            layout: IMekkoColumnLayout,
            itemCS: ClassAndSelector,
            hasSelection: boolean): UpdateSelection<MekkoChartColumnDataPoint> {

            const dataSelector: (dataPoint: MekkoChartSeries) => any[] =
                (dataPoint: MekkoChartSeries) => dataPoint.data;


            const shapeSelection: UpdateSelection<any> = series.selectAll(itemCS.selector),
                shapes: UpdateSelection<MekkoChartColumnDataPoint> = shapeSelection.data(
                    dataSelector,
                    (dataPoint: MekkoChartColumnDataPoint) => dataPoint.key);

            shapes
                .enter()
                .append("rect")
                .attr("class", (dataPoint: MekkoChartColumnDataPoint) => {
                    return itemCS.class.concat(dataPoint.highlight
                        ? " highlight"
                        : "");
                });

            shapes
                .style({
                    "fill": (dataPoint: MekkoChartColumnDataPoint) => data.showAllDataPoints
                        ? dataPoint.color
                        : data.defaultDataPointColor,
                    "fill-opacity": (dataPoint: MekkoChartColumnDataPoint) => utils.getFillOpacity(
                        dataPoint.selected,
                        dataPoint.highlight,
                        hasSelection,
                        data.hasHighlights)
                })
                .attr(layout.shapeLayout as any);

            shapes
                .exit()
                .remove();

            const borderSelection: UpdateSelection<any> = series.selectAll(BaseVisualStrategy.BorderSelector.selector),
                borders: UpdateSelection<MekkoChartColumnDataPoint> = borderSelection.data(
                    dataSelector,
                    (dataPoint: MekkoChartColumnDataPoint) => dataPoint.key);

            const borderColor: string = columnChart.BaseColumnChart.getBorderColor(data.borderSettings);

            borders
                .enter()
                .append("rect")
                .classed(BaseVisualStrategy.BorderSelector.class, true);

            borders
                .style({
                    "fill": borderColor,
                    "fill-opacity": (dataPoint: MekkoChartColumnDataPoint) => {
                        return data.hasHighlights
                            ? utils.DimmedOpacity
                            : utils.DefaultOpacity;
                    }
                })
                .attr(layout.shapeBorder as any);

            borders
                .exit()
                .remove();

            return shapes;
        }

        public selectColumn(selectedColumnIndex: number, lastSelectedColumnIndex: number): void {
            utils.setChosenColumnOpacity(
                this.graphicsContext.mainGraphicsContext,
                BaseVisualStrategy.ItemSelector.selector,
                selectedColumnIndex,
                lastSelectedColumnIndex);

            this.moveHandle(selectedColumnIndex);
        }

        public getClosestColumnIndex(x: number): number {
            return utils.getClosestColumnIndex(x, this.getColumnsCenters());
        }

        /**
         * Get the chart's columns centers (x value).
         */
        private getColumnsCenters(): number[] {
            if (!this.columnsCenters) {
                const categoryWidth: number = this.categoryLayout.categoryThickness
                    * (BaseVisualStrategy.DefaultInnerPaddingRatio - MekkoChart.InnerPaddingRatio);

                if (this.data.series.length > 0) {
                    let xScaleOffset: number = 0;

                    if (!this.categoryLayout.isScalar) {
                        xScaleOffset = categoryWidth / BaseVisualStrategy.CategoryWidthDelimiter;
                    }

                    const firstSeries: MekkoChartSeries = this.data.series[0];

                    if (firstSeries && firstSeries.data) {
                        this.columnsCenters = firstSeries.data.map((dataPoint: MekkoChartColumnDataPoint) => {
                            const value: number = this.categoryLayout.isScalar
                                ? dataPoint.categoryValue
                                : dataPoint.categoryIndex;

                            return this.xProps.scale(value) + xScaleOffset;
                        });
                    }
                }
            }

            return this.columnsCenters;
        }

        private moveHandle(selectedColumnIndex: number) {
            const columnCenters: number[] = this.getColumnsCenters(),
                x: number = columnCenters[selectedColumnIndex];

            if (!this.columnSelectionLineHandle) {
                const handleSelection: Selection<any> = this.graphicsContext.mainGraphicsContext.append("g");

                this.columnSelectionLineHandle = handleSelection;

                handleSelection
                    .append("line")
                    .classed(BaseVisualStrategy.InteractiveHoverLineSelector.class, true)
                    .attr({
                        x1: x,
                        x2: x,
                        y1: 0,
                        y2: this.height,
                    });

                handleSelection
                    .append("circle")
                    .attr({
                        cx: x,
                        cy: this.height,
                        r: PixelConverter.toString(BaseVisualStrategy.CircleRadius)
                    })
                    .classed(BaseVisualStrategy.DragHandleSelector.class, true);
            }
            else {
                const handleSelection: Selection<any> = this.columnSelectionLineHandle;

                handleSelection
                    .select("line")
                    .attr({
                        x1: x,
                        x2: x
                    });

                handleSelection
                    .select("circle")
                    .attr({ cx: x });
            }
        }

        public static getLayout(
            data: MekkoColumnChartData,
            axisOptions: MekkoColumnAxisOptions): IMekkoColumnLayout {

            const xScale: LinearScale<number, number> = axisOptions.xScale,
                yScale: LinearScale<number, number> = axisOptions.yScale,
                scaledY0: number = yScale(0),
                scaledX0: number = xScale(0),
                borderWidth: number = columnChart.BaseColumnChart.getBorderWidth(data.borderSettings);

            const columnWidthScale: LayoutFunction = (dataPoint: MekkoChartColumnDataPoint) => {
                return AxisHelper.diffScaled(xScale, dataPoint.value, 0);
            };

            const columnStart: LayoutFunction = (dataPoint: MekkoChartColumnDataPoint) => {
                return scaledX0
                    + AxisHelper.diffScaled(xScale, dataPoint.originalPosition, 0)
                    + borderWidth * dataPoint.categoryIndex;
            };

            const borderStart: LayoutFunction = (dataPoint: MekkoChartColumnDataPoint) => {
                return scaledX0
                    + AxisHelper.diffScaled(xScale, dataPoint.originalPosition, 0)
                    + AxisHelper.diffScaled(xScale, dataPoint.value, 0)
                    + borderWidth * dataPoint.categoryIndex;
            };

            const yPosition: LayoutFunction = (dataPoint: MekkoChartColumnDataPoint) => {
                return scaledY0 + AxisHelper.diffScaled(yScale, dataPoint.position, 0);
            };

            const height: LayoutFunction = (dataPoint: MekkoChartColumnDataPoint) => {
                return utils.getSize(yScale, dataPoint.valueAbsolute);
            };

            return {
                shapeLayout: {
                    width: columnWidthScale,
                    x: columnStart,
                    y: yPosition,
                    height: height
                },
                shapeBorder: {
                    width: () => borderWidth,
                    x: borderStart,
                    y: yPosition,
                    height: height
                },
                shapeLayoutWithoutHighlights: {
                    width: columnWidthScale,
                    x: columnStart,
                    y: yPosition,
                    height: (dataPoint: MekkoChartColumnDataPoint) => {
                        return utils.getSize(yScale, dataPoint.originalValueAbsolute);
                    }
                },
                zeroShapeLayout: {
                    width: columnWidthScale,
                    x: columnStart,
                    y: (dataPoint: MekkoChartColumnDataPoint) => {
                        return scaledY0 + AxisHelper.diffScaled(yScale, dataPoint.position, 0)
                            + utils.getSize(yScale, dataPoint.valueAbsolute);
                    },
                    height: () => 0
                },
                shapeXAxis: {
                    width: columnWidthScale,
                    x: columnStart,
                    y: yPosition,
                    height: height
                },
            };
        }

        private createMekkoLabelDataPoints(): LabelDataPoint[] {
            const labelDataPoints: LabelDataPoint[] = [],
                data: MekkoChartData = this.data,
                dataSeries: MekkoChartSeries[] = data.series,
                formattersCache: IColumnFormatterCacheManager = createColumnFormatterCacheManager(),
                shapeLayout = this.layout.shapeLayout;

            for (let currentSeries of dataSeries) {
                const labelSettings: VisualDataLabelsSettings = currentSeries.labelSettings
                    ? currentSeries.labelSettings
                    : data.labelSettings;

                if (!labelSettings.show || !currentSeries.data) {
                    continue;
                }

                const displayUnitValue: number = getDisplayUnitValueFromAxisFormatter(
                    this.yProps.formatter,
                    labelSettings);

                for (let dataPoint of currentSeries.data) {
                    if ((data.hasHighlights && !dataPoint.highlight)
                        || dataPoint.value == null) {
                        continue;
                    }

                    const parentRect: IRect = {
                        left: shapeLayout.x(dataPoint),
                        top: shapeLayout.y(dataPoint),
                        width: shapeLayout.width(dataPoint),
                        height: shapeLayout.height(dataPoint),
                    };

                    let formatString: string = null,
                        value: number = dataPoint.valueOriginal;

                    if (!labelSettings.displayUnits) {
                        formatString = hundredPercentFormat;
                        value = dataPoint.valueAbsolute;
                    }

                    const formatter: IValueFormatter = formattersCache.getOrCreate(
                        formatString,
                        labelSettings,
                        displayUnitValue);

                    labelDataPoints.push({
                        parentRect,
                        text: formatter.format(value),
                        fillColor: labelSettings.labelColor
                            ? labelSettings.labelColor
                            : BaseVisualStrategy.DefaultLabelFillColor
                    });
                }
            }

            return labelDataPoints;
        }
    }

    export function getDisplayUnitValueFromAxisFormatter(
        axisFormatter: IValueFormatter,
        labelSettings: VisualDataLabelsSettings): number {

        if (axisFormatter
            && axisFormatter.displayUnit
            && labelSettings.displayUnits === 0) {

            return axisFormatter.displayUnit.value;
        }

        return null;
    }
}
