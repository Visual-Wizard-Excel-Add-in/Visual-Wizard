declare const OfficeRuntime: typeof import("office-runtime");

interface MessageType {
  type: string;
  title: string;
  body: string;
}

interface GraphNodeType {
  address: string;
  condition?: string;
  dependencies?: GraphNodeType[];
  falseValue?: string;
  formula: string;
  functionName?: string;
  trueValue?: string;
  values?: string[];
  criteriaRange?: string;
  criteria?: string | string[];
  sumRange?: string;
  conditions?: string[];
  criteriaRanges?: string[];
}

interface GraphType {
  data: GraphNodeType;
  dependencies: Set<GraphType>;
}

type EdgesType = ("EdgeBottom" | "EdgeLeft" | "EdgeTop" | "EdgeRight")[];

interface BorderStyleType {
  [key: string]: {
    color: string;
    style:
      | Excel.BorderLineStyle
      | "None"
      | "Continuous"
      | "Dash"
      | "DashDot"
      | "DashDotDot"
      | "Dot"
      | "Double"
      | "SlantDashDot";
    weight: Excel.BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick";
  };
}

interface StylePresetsType {
  [key: string]: CellStyleType;
}

interface CellStyleType {
  borders: BorderStyleType;
  numberFormat: string;
  numberFormatLocal?: string;
  font: {
    color: string;
    bold: boolean;
    size: number;
    italic: boolean;
    underline:
      | Excel.RangeUnderlineStyle
      | "None"
      | "Single"
      | "Double"
      | "SingleAccountant"
      | "DoubleAccountant";
    name: string;
    tintAndShade?: number;
    strikethrough?: boolean;
  };
  fill: {
    color: string;
  };
  alignment: {
    horizontalAlignment:
      | Excel.HorizontalAlignment
      | "General"
      | "Left"
      | "Center"
      | "Right"
      | "Fill"
      | "Justify"
      | "CenterAcrossSelection"
      | "Distributed";
    verticalAlignment:
      | Excel.VerticalAlignment
      | "Top"
      | "Center"
      | "Bottom"
      | "Justify"
      | "Distributed";
    wrapText: boolean;
    indentLevel: number;
    readingOrder:
      | Excel.ReadingOrder
      | "Context"
      | "LeftToRight"
      | "RightToLeft";
    textOrientation: number;
  };
  protection: {
    locked: boolean;
    formulaHidden: boolean;
  };
}

type BorderLineType =
  | "None"
  | "Continuous"
  | "Dash"
  | "DashDot"
  | "Dot"
  | "Automatic"
  | ChartLineStyle
  | "Grey25"
  | "Grey50"
  | "Grey75"
  | "RoundDot";

type ChartType =
  | "Invalid"
  | "ColumnClustered"
  | "ColumnStacked"
  | "ColumnStacked100"
  | "3DColumnClustered"
  | "3DColumnStacked"
  | "3DColumnStacked100"
  | "BarClustered"
  | "BarStacked"
  | "BarStacked100"
  | "3DBarClustered"
  | "3DBarStacked"
  | "3DBarStacked100"
  | "LineStacked"
  | "LineStacked100"
  | "LineMarkers"
  | "LineMarkersStacked"
  | "LineMarkersStacked100"
  | "PieOfPie"
  | "PieExploded"
  | "3DPieExploded"
  | "BarOfPie"
  | "XYScatterSmooth"
  | "XYScatterSmoothNoMarkers"
  | "XYScatterLines"
  | "XYScatterLinesNoMarkers"
  | "AreaStacked"
  | "AreaStacked100"
  | "3DAreaStacked"
  | "3DAreaStacked100"
  | "DoughnutExploded"
  | "RadarMarkers"
  | "RadarFilled"
  | "Surface"
  | "SurfaceWireframe"
  | "SurfaceTopView"
  | "SurfaceTopViewWireframe"
  | "Bubble"
  | "Bubble3DEffect"
  | "StockHLC"
  | "StockOHLC"
  | "StockVHLC"
  | "StockVOHLC"
  | "CylinderColClustered"
  | "CylinderColStacked"
  | "CylinderColStacked100"
  | "CylinderBarClustered"
  | "CylinderBarStacked"
  | "CylinderBarStacked100"
  | "CylinderCol"
  | "ConeColClustered"
  | "ConeColStacked"
  | "ConeColStacked100"
  | "ConeBarClustered"
  | "ConeBarStacked"
  | "ConeBarStacked100"
  | "ConeCol"
  | "PyramidColClustered"
  | "PyramidColStacked"
  | "PyramidColStacked100"
  | "PyramidBarClustered"
  | "PyramidBarStacked"
  | "PyramidBarStacked100"
  | "PyramidCol"
  | "3DColumn"
  | "Line"
  | "3DLine"
  | "3DPie"
  | "Pie"
  | "XYScatter"
  | "3DArea"
  | "Area"
  | "Doughnut"
  | "Radar"
  | "Histogram"
  | "Boxwhisker"
  | "Pareto"
  | "RegionMap"
  | "Treemap"
  | "Waterfall"
  | "Sunburst"
  | "Funnel";

interface ChartStyleType {
  chartType: ChartType;
  font: {
    name: string;
    size: number;
    color: string;
    bold: boolean;
    italic: boolean;
    underline: "None" | "Single" | Excel.ChartUnderlineStyle;
  };
  roundedCorners: boolean;
  fill: {
    color: stirng;
  };
  border: {
    lineStyle: BorderLineType;
    color: string;
    weight: number;
  };
  plotArea: {
    fill: string;
    border: {
      lineStyle: BorderLineType;
      color: string;
      weight: number;
    };
    position: "Automatic" | Excel.ChartPlotAreaPosition | "Custom";
    height: number;
    left: number;
    top: number;
    width: number;
    insideHeight: number;
    insideLeft: number;
    insideTop: number;
    insideWidth: number;
  };
  legend: {
    fill: string;
    font: {
      name: string;
      size: number;
      color: string;
      bold: boolean;
      italic: boolean;
      underline: "None" | "Single" | Excel.ChartUnderlineStyle;
    };
    border: {
      lineStyle: BorderLineType;
      color: string;
      weight: number;
    };
    position:
      | "Invalid"
      | "Custom"
      | "Left"
      | "Right"
      | "Top"
      | "Bottom"
      | Excel.ChartLegendPosition
      | "Corner";
  };
  seriesNameLevel: number;
  axes: {
    categoryAxis?: {
      position:
        | "Automatic"
        | "Custom"
        | Excel.ChartAxisPosition
        | "Maximum"
        | "Minimum";
      format: {
        line: {
          color: string;
          style: BorderLineType;
          weight: number;
        };
        font: {
          name: string;
          size: number;
          color: string;
          bold: boolean;
          italic: boolean;
          underline: "None" | "Single" | Excel.ChartUnderlineStyle;
        };
      };
    };
    valueAxis?: {
      position:
        | "Automatic"
        | "Custom"
        | Excel.ChartAxisPosition
        | "Maximum"
        | "Minimum";
      format: {
        line: {
          color: string;
          style: BorderLineType;
          weight: number;
        };
        font: {
          name: string;
          size: number;
          color: string;
          bold: boolean;
          italic: boolean;
          underline: "None" | "Single" | Excel.ChartUnderlineStyle;
        };
      };
    };
  };
  series: Excel.ChartSeries[];
}

interface ValueAxisType {
  position:
    | "Automatic"
    | "Custom"
    | Excel.ChartAxisPosition
    | "Maximum"
    | "Minimum";
  format: {
    line: {
      color: string;
      style: BorderLineType;
      weight: number;
    };
    font: {
      name: string;
      size: number;
      color: string;
      bold: boolean;
      italic: boolean;
      underline: "None" | "Single" | Excel.ChartUnderlineStyle;
    };
  };
}

interface MacroActionType {
  type:
    | "WorksheetChanged"
    | "TableChanged"
    | "ChartAdded"
    | "TableAdded"
    | "WorksheetFormatChanged";
  address?: string;
  details?:
    | {
        value: any;
      }
    | Excel.ChangedEventDetail;
  cellStyle?: CellStyleType;
  chartId?: string;
  tableId?: string;
  changeType?:
    | Excel.DataChangeType
    | "Unknown"
    | "RangeEdited"
    | "RowInserted"
    | "RowDeleted"
    | "ColumnInserted"
    | "ColumnDeleted"
    | "CellInserted"
    | "CellDeleted";
  showHeaders?: boolean;
  chartType?: ChartType;
  position?: {
    top: number;
    left: number;
  };
  size?: {
    height: number;
    width: number;
  };
  dataRange?: string[];
}

interface MacroPresetsType {
  [key: string]: { actions: MacroActionType[] };
}
