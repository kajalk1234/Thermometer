/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ''Software''), to deal
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
module powerbi.extensibility.visual {
    import Selection = d3.Selection;
    import DataView = powerbi.DataView;
    import IViewport = powerbi.IViewport;
    import VisualObjectInstance = powerbi.VisualObjectInstance;
    import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
    import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
    import createLegend = powerbi.extensibility.utils.chart.legend.createLegend;
    import ILegend = powerbi.extensibility.utils.chart.legend.ILegend;
    import legend = powerbi.extensibility.utils.chart.legend;
    import LegendData = powerbi.extensibility.utils.chart.legend.LegendData;
    import LegendDataPoint = powerbi.extensibility.utils.chart.legend.LegendDataPoint;
    import LegendIcon = powerbi.extensibility.utils.chart.legend.LegendIcon;
    import LegendPosition = powerbi.extensibility.utils.chart.legend.LegendPosition;
    import IInteractivityService = powerbi.extensibility.utils.interactivity.IInteractivityService;
    import createInteractivityService = powerbi.extensibility.utils.interactivity.createInteractivityService;
    import DataViewObjectsParser = powerbi.extensibility.utils.dataview.DataViewObjectsParser;
    import TextProperties = powerbi.extensibility.utils.formatting.TextProperties;
    import textMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;
    import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;

    export class ThermometerSettings extends DataViewObjectsParser {
        public legend: LegendSettings = new LegendSettings();
    }

    export class LegendSettings {
        public show: boolean = true;
        public position: string = 'Right';
        public showTitle: boolean = false;
        public titleText: string = '';
        public labelColor: string = '#000000';
        public fontSize: number = 8;
    }

    export interface IViewModel {
        value: number;
        targetValue: number;
        color?: string;
        min?: number;
        max?: number;
        drawTickBar?: boolean;
    }
    export interface IRange {
        range1: number;
        range2: number;
        range3: number;
        range4: number;
    }

    export interface IThermometerViewModel {
        dataView: DataView;
        settings: ThermometerSettings;
        legendData: LegendData;
    }

    export class Thermometer implements IVisual {
        private viewport: IViewport;
        private events: IVisualEventService;
        // tslint:disable-next-line:no-any
        private settings : DataViewObjectsParser;
        private svg: d3.Selection<SVGElement>;
        private mainGroup: d3.Selection<SVGElement>;
        private backCircle: d3.Selection<SVGElement>;
        private backRect: d3.Selection<SVGElement>;
        private fillCircle: d3.Selection<SVGElement>;
        private fillRect: d3.Selection<SVGElement>;
        private tempMarkings: d3.Selection<SVGElement>;
        private target: d3.Selection<SVGElement>;
        private text: d3.Selection<SVGElement>;
        private data: IViewModel;
        private range: IRange;
        private dataView: DataView;
        private viewModel: IThermometerViewModel;
        private fill: string;
        private border: string;
        private interactivityService: IInteractivityService;
        private host: IVisualHost;
        private legend: ILegend;
        private body: Selection<{}>;
        // tslint:disable-next-line:no-any
        private legendsTitleData : any = [];
        // tslint:disable-next-line:no-any
        private valFormatter : any;
        // tslint:disable-next-line:no-any
        private legendsFormatter : any ;
        // tslint:disable-next-line:no-any
        private h1 : any; private h2 : any; private h3 : any; private h4 : any;

        // This is called once when the visual is initialially created
        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            const svg : d3.Selection<SVGAElement> = this.svg = d3.select(options.element)
            .style('height', '100%').append('svg').classed('Thermometer', true);
            options.element.setAttribute('id', 'container');
            this.body = d3.select(options.element);
            this.interactivityService = createInteractivityService(this.host);
            this.body.style('cursor', 'default');
            this.legend = createLegend(
                options.element,
                false,
                null,
                true);
            this.events = options.host.eventService;
        }

        private static getValue<T>(dataView: DataView, objectName: string, key: string, defaultValue: T): T {
            if (dataView) {
                const objects : DataViewObjects = dataView.metadata.objects;
                if (objects) {
                    const config : DataViewObject = objects[objectName];
                    if (config) {
                        if (config[`valDecimalValue`] > 4) {
                            config[`valDecimalValue`] = 4;
                        } else if (config[`valDecimalValue`] < 0) {
                            config[`valDecimalValue`] = 0;
                        }
                        if (config[`decimalValue`] > 4) {
                            config[`decimalValue`] = 4;
                        } else if (config[`decimalValue`] < 0) {
                            config[`decimalValue`] = 0;
                        }
                        const val : T = <T>config[key];
                        if (val != null) {
                            return val;
                        }
                    }
                }
            }

            return defaultValue;
        }

        // tslint:disable-next-line:no-any
        private static getFill(dataView: DataView, key : any): Fill {
            if (dataView) {
                const objects : DataViewObject = dataView.metadata.objects;

                if (objects) {
                    const config : DataViewPropertyValue = objects[`config`];
                    if (config) {
                        const fill : Fill = <Fill>config[key];
                        if (fill) {
                            return fill;
                        }
                    }
                    const category1 :  DataViewPropertyValue = objects[`ranges`];
                    if (category1) {
                        const fill : Fill = <Fill>category1[key];
                        if (fill) {
                            return fill;
                        }
                    }
                    const legend : DataViewPropertyValue = objects[`legend`];
                    if (legend) {
                        const fill : Fill = <Fill>legend[key];
                        if (fill) {
                            return fill;
                        }
                    }
                }
            }
            switch (key) {
                case 'fontColor':
                    return { solid: { color: '#000' } };
                case 'fill1':
                    return { solid: { color: '#13D9F9' } };
                case 'fill2':
                    return { solid: { color: '#86EA22' } };
                case 'fill3':
                    return { solid: { color: '#FBA91C' } };
                case 'fill4':
                    return { solid: { color: '#EC4427' } };
                case 'border1':
                    return { solid: { color: '#D0EEF7' } };
                case 'border2':
                    return { solid: { color: '#D5F3B6' } };
                case 'border3':
                    return { solid: { color: '#ECCA90' } };
                case 'border4':
                    return { solid: { color: '#F9A99B' } };
                case 'labelColor':
                    return { solid: { color: '#000' } };
                default:
                    return { solid: { color: '#D0EEF7' } };
            }
        }
        // Update is called for data updates, resizes & formatting changes
        public update(options: VisualUpdateOptions) : void {            
            this.events.renderingStarted(options);
            this.viewport = options.viewport;
            if (!options.dataViews) {
                return;
            }
            if (0 === options.dataViews.length) {
                return;
            }
            const dataView : DataView = options.dataViews[0];
            this.dataView = options.dataViews[0];

            this.data = {
                value: 0,
                targetValue: null,
                color: null,
                min: null,
                max: null,
                drawTickBar: null
            };
            let tempAvailable : boolean = false;
            // tslint:disable-next-line:no-any
            const dataViewCat: any = dataView.categorical;
            if (dataView && dataViewCat) {
                for (let iCatValue : number = 0; iCatValue < dataViewCat.values.length; iCatValue++) {
                    if (options.dataViews[0].categorical.values[iCatValue].source.roles[`Temperature`]) {
                        tempAvailable = true;
                    }
                }
                if (!tempAvailable) {
                    $('.Thermometer').empty();
                    $('.legend g').empty();

                    return;
                }
                this.renderData();
                this.getFormatter(options);
                this.getRange();
                this.renderLegend();

                const viewport : IViewport = options.viewport;
                const duration : number = 1000;
                const height : number = viewport.height;
                const width : number = viewport.width;
                this.svg.attr('width', width);
                this.svg.attr('height', height);
                this.draw(width, height, duration);

                // Adding ellipsis for ticks
                const tickstarting: number = parseInt(d3.select('.y.axis').attr('transform').split('(')[1].split(',')[0], 10);
                const measureTextProperties1: TextProperties = {
                    text: d3.selectAll('.tick:last-of-type text').text(),
                    fontFamily: 'Segoe UI',
                    fontSize: `${((height * 0.1) * 0.03) * 16}px`
                };
                let tickwidth: number = 0;
                let unit = d3.selectAll('.tick:last-of-type text').text().split(" ")[1];
                let rest_digit : boolean;

                if (unit !== "" && unit.match("^\\d+$")){
                    rest_digit = true;
                }
                tickwidth = textMeasurementService.measureSvgTextWidth(measureTextProperties1);
                const thisContext : this = this;
                const ticks : d3.selection.Group = this.svg.selectAll('.tick text')[0];

                // tslint:disable-next-line:typedef
                ticks.forEach(element =>  {
                    // tslint:disable-next-line:no-any
                    let ele : any;
                    ele = element;
                    if (rest_digit){
                        ele.textContent = thisContext.getTMS(
                            ele.textContent.split(" ")[0], ((height * 0.1) * 0.03) * 16,
                            tickwidth + (options.viewport.width - (tickwidth + tickstarting + 9)));
                    }
                    else{
                        ele.textContent = thisContext.getTMS(
                            ele.textContent, ((height * 0.1) * 0.03) * 16,
                            tickwidth + (options.viewport.width - (tickwidth + tickstarting + 9)));
                    }
                });

                // Adding title for legends
                d3.selectAll('.legendItem > title')
                    // tslint:disable-next-line:no-any
                    .text( (d:any, i:number)=> {
                        return thisContext.legendsTitleData[i];
                    });
            }
            this.events.renderingFinished(options);
        }
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            const instances: VisualObjectInstance[] = [];
            const dataView : DataView = this.dataView;

            switch (options.objectName) {
                case 'config':
                    const config: VisualObjectInstance = {
                        objectName: 'config',
                        displayName: 'Configurations',
                        selector: null,
                        properties: {
                            max: Thermometer.getValue<number>(dataView, 'config', 'max', null),
                            min: Thermometer.getValue<number>(dataView, 'config', 'min', null),
                            tickBar: Thermometer.getValue<boolean>(dataView, 'config', 'tickBar', true),
                            fontColor: Thermometer.getFill(dataView, 'fontColor'),
                            postfix: Thermometer.getValue<string>(dataView, 'config', 'postfix', ''),
                            valDisplayUnits: Thermometer.getValue(dataView, 'config', 'valDisplayUnits', 0),
                            valDecimalValue: Thermometer.getValue(dataView, 'config', 'valDecimalValue', 0)
                        }
                    };
                    instances.push(config);
                    break;
                case 'legend':
                    const legend: VisualObjectInstance = {
                        objectName: 'legend',
                        displayName: 'Legend',
                        selector: null,
                        properties: {
                            show: Thermometer.getValue<boolean>(dataView, 'legend', 'show', true),
                            position: Thermometer.getValue<string>(dataView, 'legend', 'position', 'Right'),
                            showTitle: Thermometer.getValue<boolean>(dataView, 'legend', 'showTitle', false),
                            titleText: Thermometer.getValue<string>(dataView, 'legend', 'titleText', null),
                            labelColor: Thermometer.getFill(dataView, 'labelColor'),
                            displayUnits: Thermometer.getValue(dataView, 'legend', 'displayUnits', 0),
                            decimalValue: Thermometer.getValue(dataView, 'legend', 'decimalValue', 0),
                            range1: this.range.range1,
                            fill1: Thermometer.getFill(dataView, 'fill1'),
                            border1: Thermometer.getFill(dataView, 'border1'),
                            range2: this.range.range2,
                            fill2: Thermometer.getFill(dataView, 'fill2'),
                            border2: Thermometer.getFill(dataView, 'border2'),
                            range3: this.range.range3,
                            fill3: Thermometer.getFill(dataView, 'fill3'),
                            border3: Thermometer.getFill(dataView, 'border3'),
                            range4: this.data.max,
                            fill4: Thermometer.getFill(dataView, 'fill4'),
                            border4: Thermometer.getFill(dataView, 'border4')
                        }
                    };
                    instances.push(legend);
                    break;
                default:
                    break;
            }

            return instances;
        }
        private getFormatter(options: VisualUpdateOptions) : void {
            // Formatter for ticks
            let displayVal : number = 0;
            if (Thermometer.getValue(this.dataView, 'config', 'valDisplayUnits', 0) === 0) {
                const valLen : number = this.data.max.toString().length;
                if (valLen > 9) {
                    displayVal = 1e9;
                } else if (valLen <= 9 && valLen > 6) {
                    displayVal = 1e6;
                } else if (valLen <= 6 && valLen >= 4) {
                    displayVal = 1e3;
                } else {
                    displayVal = 10;
                }
            }
            if (options.dataViews[0].categorical.values[0].source.format &&
                options.dataViews[0].categorical.values[0].source.format.indexOf('%') !== -1) {
                this.valFormatter = valueFormatter.create({
                    format: this.dataView.categorical.values[0].source.format,
                    value: Thermometer.getValue(this.dataView, 'config', 'valDisplayUnits', 0) === 0 ?
                        0 : Thermometer.getValue(this.dataView, 'config', 'valDisplayUnits', 0),
                    precision: Thermometer.getValue(this.dataView, 'config', 'valDecimalValue', 0)
                });
            } else {
                this.valFormatter = valueFormatter.create({
                    format: this.dataView.categorical.values[0].source.format,
                    value: Thermometer.getValue(this.dataView, 'config', 'valDisplayUnits', 0) === 0 ?
                        displayVal : Thermometer.getValue(this.dataView, 'config', 'valDisplayUnits', 0),
                    precision: Thermometer.getValue(this.dataView, 'config', 'valDecimalValue', 0)
                });
            }
            if (options.dataViews[0].categorical.values[0].source.format &&
                options.dataViews[0].categorical.values[0].source.format.indexOf('%') !== -1) {
                this.legendsFormatter = valueFormatter.create({
                    format: this.dataView.categorical.values[0].source.format,
                    value: Thermometer.getValue(this.dataView, 'legend', 'displayUnits', 0) === 0 ?
                        0 : Thermometer.getValue(this.dataView, 'legend', 'displayUnits', 0),
                    precision: Thermometer.getValue(this.dataView, 'legend', 'decimalValue', 0)
                });
            } else {
                this.legendsFormatter = valueFormatter.create({
                    format: this.dataView.categorical.values[0].source.format,
                    value: Thermometer.getValue(this.dataView, 'legend', 'displayUnits', 0) === 0 ?
                        displayVal : Thermometer.getValue(this.dataView, 'legend', 'displayUnits', 0),
                    precision: Thermometer.getValue(this.dataView, 'legend', 'decimalValue', 0)
                });
            }
        }
        private renderData() : void {
            const series : DataViewValueColumns = this.dataView.categorical.values;

            if (series && 0 !== series.length) {
                const length : number = series.length;
                let iCount : number = 0;
                while (iCount < length) {
                    if (series[iCount].source.roles[`Temperature`]) {
                        if (typeof series[iCount].values[0] === 'number') {
                            this.data.value = <number>series[iCount].values[0];
                        }
                    }
                    if (series[iCount].source.roles[`TargetValue`]) {
                        if (typeof series[iCount].values[0] === 'number') {
                            this.data.targetValue = <number>series[iCount].values[0];
                        }
                    }
                    if (series[iCount].source.roles[`Min`]) {
                        if (typeof series[iCount].values[0] === 'number') {
                            this.data.min = <number>series[iCount].values[0];
                        }
                    }
                    if (series[iCount].source.roles[`Max`]) {
                        if (typeof series[iCount].values[0] === 'number') {
                            this.data.max = <number>series[iCount].values[0];
                        }
                    }
                    iCount++;
                }
            }

            if (!this.data.max) {
                this.data.max = Thermometer.getValue(this.dataView, 'config', 'max', 100);
            }
            if (!this.data.min) {
                this.data.min = Thermometer.getValue(this.dataView, 'config', 'min', 0);
            }

            this.data.drawTickBar = Thermometer.getValue(this.dataView, 'config', 'tickBar', true);
            let maximum : number = this.data.max;
            let minimum : number = this.data.min;
            if (this.data.targetValue) {
                maximum = Math.max(this.data.value, this.data.targetValue);
                minimum = Math.min(this.data.value, this.data.targetValue);
            } else {
                maximum = Math.max(this.data.value, maximum);
                minimum = Math.min(this.data.value, minimum);
            }
            // to handle value greater than max value
            if (maximum > this.data.max) {
                this.data.max = Math.ceil(maximum);
            }
            if (minimum < this.data.min) {
                this.data.min = Math.floor(minimum);
            }
            if (this.data.min >= this.data.max) {
                this.data.min = this.data.max - 1;
            }

        }
        private getRange() : void {
            this.range = {
                range1: 0,
                range2: 0,
                range3: 0,
                range4: 0
            };

            this.range.range1 = Thermometer.getValue(this.dataView, 'legend', 'range1', null);
            if (this.range.range1 !== null && this.range.range1 >= this.data.max || this.range.range1 < this.data.min) {
                this.range.range1 = null;
            }

            this.range.range2 = Thermometer.getValue(this.dataView, 'legend', 'range2', null);
            if (this.range.range1 === null && this.range.range2 !== null) {
                this.range.range2 = null;
            } else if (this.range.range2 <= this.range.range1 || this.range.range2 > this.data.max) {
                this.range.range2 = null;
            }

            this.range.range3 = Thermometer.getValue(this.dataView, 'legend', 'range3', null);
            if (this.range.range2 === null && this.range.range3 !== null) {
                this.range.range3 = null;
            } else if (this.range.range3 <= this.range.range2 || this.range.range3 > this.data.max) {
                this.range.range3 = null;
            }

            // Category-4, being the last category, always ends at maximum value
            this.range.range4 = this.data.max;
            this.getRangeColor();
        }

        private getRangeColor() : void {
            const settings: ThermometerSettings = ThermometerSettings.parse<ThermometerSettings>(this.dataView);
            const legendData : legend.LegendData = this.createLegendData(this.dataView, this.host, settings);
            this.viewModel = {
                dataView: this.dataView,
                settings: settings,
                legendData: legendData
            };

            for (let iterator : number = 0; iterator < legendData.dataPoints.length; iterator++) {
                if (iterator === 0) {
                    if (this.data.value >= this.data.min && this.data.value <= this.h1) {
                        this.fill = 'fill1';
                        this.border = 'border1';
                    }
                } else if (iterator === 1) {
                    if (this.data.value > this.h1 && this.data.value <= this.h2) {
                        this.fill = 'fill2';
                        this.border = 'border2';
                    }
                } else if (iterator === 2) {
                    if (this.data.value > this.h2 && this.data.value <= this.h3) {
                        this.fill = 'fill3';
                        this.border = 'border3';
                    }
                } else if (iterator === 3) {
                    if (this.data.value > this.h3 && this.data.value <= this.h4) {
                        this.fill = 'fill4';
                        this.border = 'border4';
                    }
                }
            }
        }
        private createLegendData(dataView: DataView, host: IVisualHost, settings: ThermometerSettings): LegendData {

            const legendData: LegendData = {
                fontSize: this.viewport.height * 0.032,
                dataPoints: [],
                title: settings.legend.showTitle ? (settings.legend.titleText) : null,
                labelColor: settings.legend.labelColor
            };

            const low : number[] = this.getLowValueForLegend();
            const high : number[]  = this.getHighValueForLegend();

            let i : number = 0;
            // tslint:disable-next-line:no-any
            const label : any[] = [];
            if (low[1] > this.data.max) {
                low[1] = this.data.max;
            }
            if (low[2] > this.data.max) {
                low[2] = this.data.max;
            }

            while (low[i] < this.data.max && i < 4) {
                label.push((i + 1).toString());
                i++;
            }
            if (low[i] === this.data.max && high[i] === this.data.max && i < 4) {
                label.push((i + 1).toString());
            }
            this.h1 = high[0];
            this.h2 = high[1];
            this.h3 = high[2];
            this.h4 = high[3];
            // let category = dataView.categorical.values[0];
            const thisContext : this = this;
            this.legendsTitleData = [];
            legendData.dataPoints = label.map(
                (typeMeta: string, index: number): LegendDataPoint => {
                    this.legendsTitleData.push(`${low[parseInt(typeMeta, 10) - 1]}${'-'}${high[parseInt(typeMeta, 10) - 1]}`);

                    return {
                        label: <string>(`${thisContext.legendsFormatter.format(low[parseInt(typeMeta, 10) - 1])}${'-'}` +
                            `${thisContext.legendsFormatter.format(high[parseInt(typeMeta, 10) - 1])}`),
                        color: Thermometer.getFill(dataView, `${'fill'}${typeMeta}`).solid.color,
                        icon: LegendIcon.Circle,
                        selected:false,
                        identity: host.createSelectionIdBuilder()
                            .withMeasure(typeMeta)
                            .createSelectionId()

                    };
                });

            return legendData;
        }

        private getLowValueForLegend(): number[] {
            // tslint:disable-next-line:no-any
            const low : any[] = [];
            for (let lowVal : number = 0; lowVal < 4; lowVal++) {
                switch (lowVal) {
                    case 0:
                    if (this.data.min !== null) {
                        low.push(this.data.min);
                    }
                    break;
                    case 1:
                if (this.range.range1 !== null) {
                        low.push(this.range.range1);
                    }
                break;

                case 2:
                    if (this.range.range2 !== null) {
                        low.push(this.range.range2);
                    }
                    break;
                case 3:
                    if (this.range.range3 !== null) {
                        low.push(this.range.range3);
                    }
                    break;
                default:
                break;
            }}

            return low;
        }
        private getHighValueForLegend(): number[] {
            // tslint:disable-next-line:no-any
            const high : any[] = [];
            for (let highVal : number = 0; highVal < 4; highVal++) {
                switch (highVal) {
                    case 0:
                    if (this.range.range1 !== null) { high.push(this.range.range1); }
                    break;

                 case 1 :
                    if (this.range.range2 !== null) { high.push(this.range.range2); }
                    break;

                case 2:
                    if (this.range.range3 !== null) { high.push(this.range.range3); }
                    break;

                case 3:
                    if (this.data.max !== null) { high.push(this.data.max); }
                    break;

                default:
                    break;
                }
            }

            return high;
        }
        private draw(width: number, height: number, duration: number) : void {
            this.svg.selectAll('*').remove();
            this.mainGroup = this.svg.append('g');
            this.backRect = this.mainGroup.append('rect');
            this.backCircle = this.mainGroup.append('circle');
            this.fillRect = this.mainGroup.append('rect');
            this.fillCircle = this.mainGroup.append('circle');
            this.text = this.mainGroup.append('text');
            this.tempMarkings = this.svg.append('g').attr('class', 'y axis');
            this.target = this.svg.append('g').attr('class', 'yLeftAxis axis');
            const radius : number = height * 0.1;
            const padding : number = radius * 0.25;
            this.drawBack(width, height, radius);
            this.drawFill(width, height, radius, padding, duration);
            this.drawTicks(width, height, radius, padding);
            this.drawText(width, height, radius, padding);
            if (this.data.targetValue) {
                this.drawTarget(width, height, radius, padding);
            }
            d3.select('#y axis').remove();
        }

        private drawBack(width: number, height: number, radius: number) : void {
            const rectHeight : number = height - radius;
            const fill : string = Thermometer.getFill(this.dataView, this.border).solid.color;
            this.backCircle
                .attr({
                    cx: width / 2,
                    cy: rectHeight,
                    r: radius
                })
                .style({
                    fill: fill
                });

            this.backRect
                .attr({
                    x: (width - radius) / 2,
                    y: 0,
                    width: radius,
                    height: rectHeight
                })
                .style({
                    fill: fill
                });
        }

        private drawFill(width: number, height: number, radius: number, padding: number, duration: number) : void {
            const innerRadius : number = radius * 0.8;
            const fillWidth : number = innerRadius * 0.7;
            const zeroValue : number = height - (radius * 2) - padding;
            const fill : string = Thermometer.getFill(this.dataView, this.fill).solid.color;

            const min : number = this.data.min;
            const max : number = this.data.max;

            const value : number = this.data.value;
            let percentage : number = (zeroValue - padding) * ((value - min) / (max - min));
            let rectHeight : number = height - radius;
            if (isNaN(rectHeight)) {
                rectHeight = 0;
            }
            if (isNaN(percentage)) {
                percentage = 0;
            }
            this.fillCircle.attr({
                cx: width / 2,
                cy: rectHeight,
                r: innerRadius
            }).style({
                fill: fill
            });
            if (rectHeight !== 0 && percentage !== 0) {
                this.fillRect
                    .style({
                        fill: fill
                    })
                    .attr({
                        x: (width - fillWidth) / 2,
                        width: fillWidth
                    })
                    .attr({
                        y: zeroValue - percentage,
                        height: rectHeight - zeroValue + percentage
                    });
            }

        }

        private drawTicks(width: number, height: number, radius: number, padding: number) : void {
            d3.select('.y.axis').attr('visibility', 'visible');
            // tslint:disable-next-line:no-any
            let y : any;
            // tslint:disable-next-line:no-any
            let yAxis : any;
            // tslint:disable-next-line:no-any
            const tickData : any[] = [];
            let iCount : number;
            const postFix : '' = Thermometer.getValue(this.dataView, 'config', 'postfix', '');
            y = d3.scale.linear().range([height - (radius * 2) - padding, padding]);
            y.domain([this.data.min, this.data.max]);
            const interval : number = (this.data.max - this.data.min) / 5;
            tickData[0] = this.data.min;
            for (iCount = 1; iCount < 6; iCount++) {
                tickData[iCount] = tickData[iCount - 1] + interval;
            }
            yAxis = d3.svg.axis().scale(y).ticks(6).orient('right').tickValues(tickData).tickFormat(this.valFormatter.format);
            this.tempMarkings
                .attr('transform', `${'translate('}${((width + radius) / 2 + (radius * 0.15))}${',0)'}`)
                .style({
                    'font-size': `${(radius * 0.03)}em`,
                    'font-family': 'Segoe UI',
                    stroke: 'none',
                    fill: '#333'
                })
                .call(yAxis);
            this.tempMarkings.selectAll('.axis line, .axis path')
                .style({
                    stroke: '#333',
                    fill: 'none'
                });
            for (iCount = 0; iCount < document.querySelectorAll('.axis text').length; iCount++) {
                document.querySelectorAll('.axis text')[iCount].textContent =
                    `${document.querySelectorAll('.axis text')[iCount].textContent} ${postFix}`;
            }
            if (!this.data.drawTickBar) {
                d3.select('.y.axis').attr('visibility', 'hidden');
            }
            //const This = this;
            const titleFormatter : utils.formatting.IValueFormatter
            = valueFormatter.create({
                format: this.dataView.categorical.values[0].source.format
            });
            d3.selectAll('.y.axis>.tick')
                .append('title')
                // tslint:disable-next-line:no-any
                .text( (d:any)=> {
                    return `${titleFormatter.format(d)} ${postFix}`;
                });
        }

        private drawTarget(width: number, height: number, radius: number, padding: number) : void {
            let postFix : '' = Thermometer.getValue(this.dataView, 'config', 'postfix', '');
            d3.select('.yLeftAxis.axis').attr('visibility', 'visible');
            const target : number = this.data.targetValue;

            const zeroValue : number = height - (radius * 2) - padding;
            const min : number = this.data.min;
            const max : number = this.data.max;
            const percentage : number = (zeroValue - padding) * ((target - min) / (max - min));
            const yPos : number = zeroValue - percentage;
            let sText : string = `${target} ${postFix}`;
            let fTextWidth : number;
            let iTextHeight : number;
            const textProperties: TextProperties = {
                text: sText.toString(),
                fontFamily: 'Segoe UI',
                fontSize: `${(radius * 0.48)}px`
            };
            // Target information adding
            const targetFormatter : utils.formatting.IValueFormatter = valueFormatter.create({
                format: this.dataView.categorical.values[1].source.format
            });
            fTextWidth = textMeasurementService.measureSvgTextWidth(textProperties);
            iTextHeight = textMeasurementService.measureSvgTextHeight(textProperties);

            this.svg.append('line')
                .classed('targetLine', true)
                .style('stroke', 'black')
                .attr('x1', (width - radius) / 2)
                .attr('y1', yPos)
                .attr('x2', (width - radius) / 2 - (radius * 0.3))
                .attr('y2', yPos);
            
            const textPropertiesNew: TextProperties = {
                text: `${this.valFormatter.format(parseFloat(sText))} ${postFix}`,
                fontFamily: 'Segoe UI',
                fontSize: `${(radius * 0.48)}px`
            };
            if ((`${postFix}`).match("^\\d+$")){
                textPropertiesNew.text = `${this.valFormatter.format(parseFloat(sText))}`;
                postFix = '';
            }
            const fTextWidthNew : number = textMeasurementService.measureSvgTextWidth(textPropertiesNew);
            if (postFix) {
                this.drawTargetPostFix(width, height, radius, padding, fTextWidthNew, iTextHeight, sText, postFix, yPos);
            } else {
                if (this.legend.getOrientation() === LegendPosition.Left || this.legend.getOrientation() === LegendPosition.LeftCenter) {
                    const legendsWidth : number = $('.legend').width();
                    let xpos : number = (width - radius) / 2 - (radius * 0.5) - fTextWidthNew;
                    if (xpos < legendsWidth) {
                        xpos = legendsWidth;
                    }
                    this.svg.append('text')
                        .classed('targetText', true)
                        .attr({
                            x: `${xpos}px`,
                            y: `${(yPos) + iTextHeight * 0.1 + radius * 0.1}px`,
                            'font-size': `${(radius * 0.48)}px`,
                            'text-anchor': 'left'
                        })
                        .text(this.getTMS(
                            `${this.valFormatter.format(parseFloat(sText))} ${postFix}`,
                            (radius * 0.48),
                            (width / 2) - 50 - legendsWidth)
                        );
                } else {
                    let xpos : number = (width - radius) / 2 - (radius * 0.5) - fTextWidthNew;
                    if (xpos < 0) {
                        xpos = 0;
                    }
                    this.svg.append('text')
                        .classed('targetText', true)
                        .attr({
                            x: `${xpos}px`,
                            y: `${yPos + (iTextHeight * 0.1) + radius * 0.1}px`,
                            'font-size': `${(radius * 0.48)}px`,
                            'text-anchor': 'left'
                        })
                        .text(this.getTMS(
                            `${this.valFormatter.format(parseFloat(sText))} ${postFix}`,
                            (radius * 0.48),
                            (width / 2) - 50)
                        );
                }
            }

            this.svg.select('.targetText')
                .append('title')
                .text(`${targetFormatter.format(parseFloat(sText))} ${postFix}`);
        }

        private drawTargetPostFix (width: number, height: number, radius: number, padding: number, fTextWidthNew : number, iTextHeight: number, sText : string, postFix : '', yPos : number): void {
            if (this.legend.getOrientation() === LegendPosition.Left || this.legend.getOrientation() === LegendPosition.LeftCenter) {
                const legendsWidth : number = $('.legend').width();
                let xpos : number = (width - radius) / 2 - (radius * 0.5) - fTextWidthNew;
                if (xpos < legendsWidth) {
                    xpos = legendsWidth;
                }
                this.svg.append('text')
                    .classed('targetText', true)
                    .attr({
                        x: `${xpos}px`,
                        y: `${(yPos) + iTextHeight * 0.1 + radius * 0.1}px`,
                        'font-size': `${(radius * 0.48)}px`, 'text-anchor': 'left'
                    })
                    .text(this.getTMS(
                        `${this.valFormatter.format(parseFloat(sText))} ${postFix}`,
                        (radius * 0.48),
                        (width / 2) - 50 - legendsWidth)
                    );
            } else {
                let xpos : number = (width - radius) / 2 - (radius * 0.5) - fTextWidthNew;
                if (xpos < 0) {
                    xpos = 0;
                }
                this.svg.append('text')
                    .classed('targetText', true)
                    .attr({
                        x: `${xpos}px`,
                        y: `${(yPos) + iTextHeight * 0.1 + radius * 0.1}px`,
                        'font-size': `${(radius * 0.48)}px`,
                        'text-anchor': 'left'
                    })
                    .text(this.getTMS(
                        `${this.valFormatter.format(parseFloat(sText))} ${postFix}`,
                        (radius * 0.48),
                        (width / 2) - 50)
                    );
            }
        }

        private drawText(width: number, height: number, radius: number, padding: number) : void {
            this.text
                .text(this.getTMS(
                    this.valFormatter.format(
                        this.data.value > this.data.max ? this.data.max : this.data.value),
                    (radius * 0.03) * 16,
                    (radius * 0.8) * 2))
                .attr({ x: width / 2, y: height - radius, dy: '.35em' })
                .style({
                    fill: Thermometer.getFill(this.dataView, 'fontColor').solid.color,
                    'text-anchor': 'middle',
                    'font-family': 'Segoe UI',
                    'font-size': `${(radius * 0.03)}em`
                });
            const titleFormatter : utils.formatting.IValueFormatter = valueFormatter.create({
                format: this.dataView.categorical.values[0].source.format
            });
            this.text.append('title')
                .text(titleFormatter.format(this.data.value > this.data.max ? this.data.max : this.data.value));
        }

        private getTMS(stringName: string, textSize: number, width: number): string {
            const measureTextProperties: TextProperties = {
                text: stringName,
                fontFamily: 'Segoe UI',
                fontSize: `${textSize}px`
            };

            return textMeasurementService.getTailoredTextOrDefault(measureTextProperties, width);
        }

        /**
         * Get legend data, calculate position and draw it
         */
        private renderLegend(): void {
            if (!this.viewModel) {
                return;
            }
            if (!this.viewModel.legendData) {
                return;
            }
            const position: LegendPosition = this.viewModel.settings.legend.show
                ? LegendPosition[this.viewModel.settings.legend.position]
                : LegendPosition.None;
            this.legend.changeOrientation(position);
            this.legend.drawLegend(this.viewModel.legendData, _.clone(this.viewport));
            this.svg.style('margin', `${0}px`);
            switch (this.legend.getOrientation()) {
                case LegendPosition.Left:
                case LegendPosition.LeftCenter:
                    break;
                case LegendPosition.Right:
                case LegendPosition.RightCenter:
                    this.viewport.width -= this.legend.getMargins().width;
                    break;
                case LegendPosition.Top:
                case LegendPosition.TopCenter:
                    this.svg.style('margin-top', `${parseInt($('.legend').css('height'), 10)}px`);
                case LegendPosition.Bottom:
                case LegendPosition.BottomCenter:
                    this.viewport.height -= 1.9 * this.legend.getMargins().height;
                    break;
                default:
                    break;
            }
        }
    }
}
