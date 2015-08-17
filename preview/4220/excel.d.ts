declare module Excel {
    class Application extends OfficeExtension.ClientObject {
        private m_calculationMode;
        public calculationMode : string;
        public calculate(calculationType: string): void;
        public handleResult(value: any): void;
    }
    class Workbook extends OfficeExtension.ClientObject {
        private m_application;
        private m_bindings;
        private m_names;
        private m_tables;
        private m_worksheets;
        public application : Application;
        public bindings : BindingCollection;
        public names : NamedItemCollection;
        public tables : TableCollection;
        public worksheets : WorksheetCollection;
        public getSelectedRange(): Range;
        public _GetObjectByReferenceId(bstrReferenceId: string): OfficeExtension.ClientResult<any>;
        public _GetObjectTypeNameByReferenceId(bstrReferenceId: string): OfficeExtension.ClientResult<string>;
        public _RemoveReference(bstrReferenceId: string): void;
        public handleResult(value: any): void;
    }
    class Worksheet extends OfficeExtension.ClientObject {
        private m_charts;
        private m_id;
        private m_index;
        private m_name;
        private m_tables;
        public charts : ChartCollection;
        public tables : TableCollection;
        public id : string;
        public index : number;
        public name : string;
        public activate(): void;
        public delete(): void;
        public getCell(row: number, column: number): Range;
        public getEntireWorksheetRange(): Range;
        public getRange(address: string): Range;
        public getUsedRange(): Range;
        public handleResult(value: any): void;
    }
    class WorksheetCollection extends OfficeExtension.ClientObject {
        private m__items;
        public items : Worksheet[];
        public add(name: string): Worksheet;
        public getActiveWorksheet(): Worksheet;
        public getItem(index: string): Worksheet;
        public handleResult(value: any): void;
    }
    class Range extends OfficeExtension.ClientObject {
        private m_address;
        private m_addressLocal;
        private m_cellCount;
        private m_columnCount;
        private m_columnIndex;
        private m_format;
        private m_formulas;
        private m_formulasLocal;
        private m_numberFormat;
        private m_rowCount;
        private m_rowIndex;
        private m_text;
        private m_values;
        private m_worksheet;
        private m__ReferenceId;
        public format : RangeFormat;
        public worksheet : Worksheet;
        public address : string;
        public addressLocal : string;
        public cellCount : number;
        public columnCount : number;
        public columnIndex : number;
        public formulas : any[][];
        public formulasLocal : any[][];
        public numberFormat : any[][];
        public rowCount : number;
        public rowIndex : number;
        public text : any[][];
        public values : any[][];
        public _ReferenceId : string;
        public clear(applyTo: string): void;
        public delete(shift: string): void;
        public getCell(row: number, column: number): Range;
        public getEntireColumn(): Range;
        public getEntireRow(): Range;
        public getUsedRange(): Range;
        public insert(shift: string): void;
        public select(): void;
        public _KeepReference(): void;
        public handleResult(value: any): void;
    }
    class NamedItemCollection extends OfficeExtension.ClientObject {
        private m__items;
        public items : NamedItem[];
        public getItem(name: string): NamedItem;
        public handleResult(value: any): void;
    }
    class NamedItem extends OfficeExtension.ClientObject {
        private m_name;
        private m_type;
        private m_value;
        private m_visible;
        private m__Id;
        public name : string;
        public type : string;
        public value : any;
        public visible : boolean;
        public _Id : string;
        public getRange(): Range;
        public handleResult(value: any): void;
    }
    class Binding extends OfficeExtension.ClientObject {
        private m_id;
        private m_type;
        public id : string;
        public type : string;
        public getRange(): Range;
        public getTable(): Table;
        public getText(): OfficeExtension.ClientResult<string>;
        public handleResult(value: any): void;
    }
    class BindingCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : Binding[];
        public count : number;
        public getItem(id: string): Binding;
        public getItemAt(index: number): Binding;
        public handleResult(value: any): void;
    }
    class TableCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : Table[];
        public count : number;
        public add(address: string, hasHeaders: boolean): Table;
        public getItem(id: any): Table;
        public getItemAt(index: number): Table;
        public handleResult(value: any): void;
    }
    class Table extends OfficeExtension.ClientObject {
        private m_columns;
        private m_id;
        private m_name;
        private m_rows;
        private m_showHeaders;
        private m_showTotals;
        private m_style;
        public columns : TableColumnCollection;
        public rows : TableRowCollection;
        public id : number;
        public name : string;
        public showHeaders : boolean;
        public showTotals : boolean;
        public style : string;
        public delete(): void;
        public getDataBodyRange(): Range;
        public getHeaderRowRange(): Range;
        public getRange(): Range;
        public getTotalRowRange(): Range;
        public handleResult(value: any): void;
    }
    class TableColumnCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : TableColumn[];
        public count : number;
        public add(index: any, values: any): TableColumn;
        public getItem(id: any): TableColumn;
        public getItemAt(index: number): TableColumn;
        public handleResult(value: any): void;
    }
    class TableColumn extends OfficeExtension.ClientObject {
        private m_id;
        private m_index;
        private m_name;
        private m_values;
        public id : number;
        public index : number;
        public name : string;
        public values : any[][];
        public delete(): void;
        public getDataBodyRange(): Range;
        public getHeaderRowRange(): Range;
        public getRange(): Range;
        public getTotalRowRange(): Range;
        public handleResult(value: any): void;
    }
    class TableRowCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : TableRow[];
        public count : number;
        public add(index: any, values: any): TableRow;
        public getItemAt(index: number): TableRow;
        public handleResult(value: any): void;
    }
    class TableRow extends OfficeExtension.ClientObject {
        private m_index;
        private m_values;
        public index : number;
        public values : any[][];
        public delete(): void;
        public getRange(): Range;
        public handleResult(value: any): void;
    }
    class RangeFormat extends OfficeExtension.ClientObject {
        private m_borders;
        private m_fill;
        private m_font;
        private m_horizontalAlignment;
        private m_verticalAlignment;
        private m_wrapText;
        public borders : RangeBorderCollection;
        public fill : RangeFill;
        public font : RangeFont;
        public horizontalAlignment : string;
        public verticalAlignment : string;
        public wrapText : boolean;
        public handleResult(value: any): void;
    }
    class RangeFill extends OfficeExtension.ClientObject {
        private m_color;
        public color : string;
        public handleResult(value: any): void;
    }
    class RangeBorder extends OfficeExtension.ClientObject {
        private m_color;
        private m_sideIndex;
        private m_style;
        private m_weight;
        public color : string;
        public sideIndex : string;
        public style : string;
        public weight : string;
        public handleResult(value: any): void;
    }
    class RangeBorderCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : RangeBorder[];
        public count : number;
        public getItem(index: string): RangeBorder;
        public getItemAt(index: number): RangeBorder;
        public handleResult(value: any): void;
    }
    class RangeFont extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_italic;
        private m_name;
        private m_size;
        private m_underline;
        public bold : boolean;
        public color : string;
        public italic : boolean;
        public name : string;
        public size : number;
        public underline : string;
        public handleResult(value: any): void;
    }
    class ChartCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : Chart[];
        public count : number;
        public add(type: string, sourceData: any, seriesBy: string): Chart;
        public getItem(name: string): Chart;
        public getItemAt(index: number): Chart;
        public _GetItem(id: string): Chart;
        public handleResult(value: any): void;
    }
    class Chart extends OfficeExtension.ClientObject {
        private m_axes;
        private m_dataLabels;
        private m_format;
        private m_height;
        private m_left;
        private m_legend;
        private m_name;
        private m_series;
        private m_title;
        private m_top;
        private m_width;
        private m__Id;
        public axes : ChartAxes;
        public dataLabels : ChartDataLabels;
        public format : ChartAreaFormat;
        public legend : ChartLegend;
        public series : ChartSeriesCollection;
        public title : ChartTitle;
        public height : number;
        public left : number;
        public name : string;
        public top : number;
        public width : number;
        public _Id : string;
        public delete(): void;
        public setData(sourceData: any, seriesBy: string): void;
        public handleResult(value: any): void;
    }
    class ChartAreaFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        public fill : ChartFill;
        public font : ChartFont;
        public handleResult(value: any): void;
    }
    class ChartSeriesCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : ChartSeries[];
        public count : number;
        public getItemAt(index: number): ChartSeries;
        public handleResult(value: any): void;
    }
    class ChartSeries extends OfficeExtension.ClientObject {
        private m_format;
        private m_name;
        private m_points;
        public format : ChartSeriesFormat;
        public points : ChartPointsCollection;
        public name : string;
        public handleResult(value: any): void;
    }
    class ChartSeriesFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_line;
        public fill : ChartFill;
        public line : ChartLineFormat;
        public handleResult(value: any): void;
    }
    class ChartPointsCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : ChartPoint[];
        public count : number;
        public getItemAt(index: number): ChartPoint;
        public handleResult(value: any): void;
    }
    class ChartPoint extends OfficeExtension.ClientObject {
        private m_format;
        public format : ChartPointFormat;
        public handleResult(value: any): void;
    }
    class ChartPointFormat extends OfficeExtension.ClientObject {
        private m_fill;
        public fill : ChartFill;
        public handleResult(value: any): void;
    }
    class ChartAxes extends OfficeExtension.ClientObject {
        private m_categoryAxis;
        private m_seriesAxis;
        private m_valueAxis;
        public categoryAxis : ChartAxis;
        public seriesAxis : ChartAxis;
        public valueAxis : ChartAxis;
        public handleResult(value: any): void;
    }
    class ChartAxis extends OfficeExtension.ClientObject {
        private m_format;
        private m_majorGridlines;
        private m_majorUnit;
        private m_maximum;
        private m_minimum;
        private m_minorGridlines;
        private m_minorUnit;
        private m_title;
        public format : ChartAxisFormat;
        public majorGridlines : ChartGridlines;
        public minorGridlines : ChartGridlines;
        public title : ChartAxisTitle;
        public majorUnit : any;
        public maximum : any;
        public minimum : any;
        public minorUnit : any;
        public handleResult(value: any): void;
    }
    class ChartAxisFormat extends OfficeExtension.ClientObject {
        private m_font;
        private m_line;
        public font : ChartFont;
        public line : ChartLineFormat;
        public handleResult(value: any): void;
    }
    class ChartAxisTitle extends OfficeExtension.ClientObject {
        private m_format;
        private m_text;
        private m_visible;
        public format : ChartAxisTitleFormat;
        public text : string;
        public visible : boolean;
        public handleResult(value: any): void;
    }
    class ChartAxisTitleFormat extends OfficeExtension.ClientObject {
        private m_font;
        public font : ChartFont;
        public handleResult(value: any): void;
    }
    class ChartDataLabels extends OfficeExtension.ClientObject {
        private m_format;
        private m_position;
        private m_separator;
        private m_showBubbleSize;
        private m_showCategoryName;
        private m_showLegendKey;
        private m_showPercentage;
        private m_showSeriesName;
        private m_showValue;
        public format : ChartDataLabelFormat;
        public position : string;
        public separator : string;
        public showBubbleSize : boolean;
        public showCategoryName : boolean;
        public showLegendKey : boolean;
        public showPercentage : boolean;
        public showSeriesName : boolean;
        public showValue : boolean;
        public handleResult(value: any): void;
    }
    class ChartDataLabelFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        public fill : ChartFill;
        public font : ChartFont;
        public handleResult(value: any): void;
    }
    class ChartGridlines extends OfficeExtension.ClientObject {
        private m_format;
        private m_visible;
        public format : ChartGridlinesFormat;
        public visible : boolean;
        public handleResult(value: any): void;
    }
    class ChartGridlinesFormat extends OfficeExtension.ClientObject {
        private m_line;
        public line : ChartLineFormat;
        public handleResult(value: any): void;
    }
    class ChartLegend extends OfficeExtension.ClientObject {
        private m_format;
        private m_overlay;
        private m_position;
        private m_visible;
        public format : ChartLegendFormat;
        public overlay : boolean;
        public position : string;
        public visible : boolean;
        public handleResult(value: any): void;
    }
    class ChartLegendFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        public fill : ChartFill;
        public font : ChartFont;
        public handleResult(value: any): void;
    }
    class ChartTitle extends OfficeExtension.ClientObject {
        private m_format;
        private m_overlay;
        private m_text;
        private m_visible;
        public format : ChartTitleFormat;
        public overlay : boolean;
        public text : string;
        public visible : boolean;
        public handleResult(value: any): void;
    }
    class ChartTitleFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        public fill : ChartFill;
        public font : ChartFont;
        public handleResult(value: any): void;
    }
    class ChartFill extends OfficeExtension.ClientObject {
        public clear(): void;
        public setSolidColor(color: string): void;
        public handleResult(value: any): void;
    }
    class ChartLineFormat extends OfficeExtension.ClientObject {
        private m_color;
        public color : string;
        public clear(): void;
        public handleResult(value: any): void;
    }
    class ChartFont extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_italic;
        private m_name;
        private m_size;
        private m_underline;
        public bold : boolean;
        public color : string;
        public italic : boolean;
        public name : string;
        public size : number;
        public underline : string;
        public handleResult(value: any): void;
    }
    class BindingType {
        static range: string;
        static table: string;
        static text: string;
    }
    class BorderIndex {
        static edgeTop: string;
        static edgeBottom: string;
        static edgeLeft: string;
        static edgeRight: string;
        static insideVertical: string;
        static insideHorizontal: string;
        static diagonalDown: string;
        static diagonalUp: string;
    }
    class BorderLineStyle {
        static none: string;
        static continuous: string;
        static dash: string;
        static dashDot: string;
        static dashDotDot: string;
        static dot: string;
        static double: string;
        static slantDashDot: string;
    }
    class BorderWeight {
        static hairline: string;
        static thin: string;
        static medium: string;
        static thick: string;
    }
    class CalculationMode {
        static automatic: string;
        static automaticExceptTables: string;
        static manual: string;
    }
    class CalculationType {
        static recalculate: string;
        static full: string;
        static fullRebuild: string;
    }
    class ClearApplyTo {
        static all: string;
        static formats: string;
        static contents: string;
    }
    class ChartDataLabelPosition {
        static invalid: string;
        static none: string;
        static center: string;
        static insideEnd: string;
        static insideBase: string;
        static outsideEnd: string;
        static left: string;
        static right: string;
        static top: string;
        static bottom: string;
        static bestFit: string;
        static callout: string;
    }
    class ChartLegendPosition {
        static invalid: string;
        static top: string;
        static bottom: string;
        static left: string;
        static right: string;
        static corner: string;
        static custom: string;
    }
    class ChartSeriesBy {
        static auto: string;
        static columns: string;
        static rows: string;
    }
    class ChartType {
        static invalid: string;
        static columnClustered: string;
        static columnStacked: string;
        static columnStacked100: string;
        static _3DColumnClustered: string;
        static _3DColumnStacked: string;
        static _3DColumnStacked100: string;
        static barClustered: string;
        static barStacked: string;
        static barStacked100: string;
        static _3DBarClustered: string;
        static _3DBarStacked: string;
        static _3DBarStacked100: string;
        static lineStacked: string;
        static lineStacked100: string;
        static lineMarkers: string;
        static lineMarkersStacked: string;
        static lineMarkersStacked100: string;
        static pieOfPie: string;
        static pieExploded: string;
        static _3DPieExploded: string;
        static barOfPie: string;
        static xyscatterSmooth: string;
        static xyscatterSmoothNoMarkers: string;
        static xyscatterLines: string;
        static xyscatterLinesNoMarkers: string;
        static areaStacked: string;
        static areaStacked100: string;
        static _3DAreaStacked: string;
        static _3DAreaStacked100: string;
        static doughnutExploded: string;
        static radarMarkers: string;
        static radarFilled: string;
        static surface: string;
        static surfaceWireframe: string;
        static surfaceTopView: string;
        static surfaceTopViewWireframe: string;
        static bubble: string;
        static bubble3DEffect: string;
        static stockHLC: string;
        static stockOHLC: string;
        static stockVHLC: string;
        static stockVOHLC: string;
        static cylinderColClustered: string;
        static cylinderColStacked: string;
        static cylinderColStacked100: string;
        static cylinderBarClustered: string;
        static cylinderBarStacked: string;
        static cylinderBarStacked100: string;
        static cylinderCol: string;
        static coneColClustered: string;
        static coneColStacked: string;
        static coneColStacked100: string;
        static coneBarClustered: string;
        static coneBarStacked: string;
        static coneBarStacked100: string;
        static coneCol: string;
        static pyramidColClustered: string;
        static pyramidColStacked: string;
        static pyramidColStacked100: string;
        static pyramidBarClustered: string;
        static pyramidBarStacked: string;
        static pyramidBarStacked100: string;
        static pyramidCol: string;
        static _3DColumn: string;
        static line: string;
        static _3DLine: string;
        static _3DPie: string;
        static pie: string;
        static xyscatter: string;
        static _3DArea: string;
        static area: string;
        static doughnut: string;
        static radar: string;
    }
    class ChartUnderlineStyle {
        static none: string;
        static single: string;
    }
    class DeleteShiftDirection {
        static up: string;
        static left: string;
    }
    class HorizontalAlignment {
        static general: string;
        static left: string;
        static center: string;
        static right: string;
        static fill: string;
        static justify: string;
        static centerAcrossSelection: string;
        static distributed: string;
    }
    class InsertShiftDirection {
        static down: string;
        static right: string;
    }
    class NamedItemType {
        static string: string;
        static integer: string;
        static double: string;
        static boolean: string;
        static range: string;
    }
    class RangeUnderlineStyle {
        static none: string;
        static single: string;
        static double: string;
        static singleAccountant: string;
        static doubleAccountant: string;
    }
    class VerticalAlignment {
        static top: string;
        static center: string;
        static bottom: string;
        static justify: string;
        static distributed: string;
    }
    class ErrorCodes {
        static accessDenied: string;
        static generalException: string;
        static insertDeleteConflict: string;
        static invalidArgument: string;
        static invalidBinding: string;
        static invalidOperation: string;
        static invalidReference: string;
        static invalidSelection: string;
        static itemNotFound: string;
        static notImplemented: string;
        static unsupportedOperation: string;
    }
}
declare module Excel {
    class ExcelClientContext extends OfficeExtension.ClientRequestContext {
        private m_workbook;
        constructor(url?: string);
        public workbook : Workbook;
    }
}
