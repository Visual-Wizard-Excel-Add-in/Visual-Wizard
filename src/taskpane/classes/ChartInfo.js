class ChartInfo {
  constructor(chartType) {
    this.type = chartType;
    this.defaultOptions = [
      "format/roundedCorners",
      "format/font/*",
      "format/border/*",
      "plotArea/format/border/*",
      "plotArea/format/*",
      "plotArea/*",
      "legend/*",
      "legend/format/*",
      "legend/format/font/*",
      "legend/format/border/*",
    ];
    this.chartStyle = null;
  }

  get loadOptions() {
    switch (this.type) {
      case "ColumnClustered":
      case "ColumnStacked":
      case "ColumnStacked100":
      case "Line":
      case "LineStacked":
      case "LineStacked100":
      case "Area":
      case "AreaStacked":
      case "AreaStacked100":
      case "Histogram":
      case "BoxWhisker":
      case "Waterfall":
      case "Funnel":
      case "3DArea":
      case "3DAreaStacked":
      case "3DAreaStacked100":
      case "3DColumn":
      case "3DColumnClustered":
      case "3DColumnStacked":
      case "3DColumnStacked100":
      case "3DLine":
      case "3DBarClustered":
      case "3DBarStacked":
      case "3DBarStacked100":
        this.defaultOptions.push(
          "axes/categoryAxis/*",
          "axes/valueAxis/*",
          "axes/categoryAxis/format/*",
          "axes/valueAxis/format/*",
          "axes/categoryAxis/format/line/*",
          "axes/categoryAxis/format/font/*",
          "axes/valueAxis/format/line/*",
          "axes/valueAxis/format/font/*",
          "series/items",
        );
        break;

      case "Pie":
      case "Doughnut":
      case "Treemap":
      case "Sunburst":
      case "3DPie":
      case "3DPieExploded":
        this.defaultOptions.push("series/items");
        break;

      case "Scatter":
      case "Bubble":
      case "Xyscatter":
      case "XyscatterLines":
      case "XyscatterLinesNoMarkers":
      case "XyscatterSmooth":
      case "XyscatterSmoothNoMarkers":
        this.defaultOptions.push(
          "axes/valueAxis/*",
          "axes/valueAxis/format/*",
          "axes/valueAxis/format/line/*",
          "axes/valueAxis/format/font/*",
          "series/items",
        );
        break;

      case "StockHLC":
      case "StockOHLC":
      case "StockVHLC":
      case "StockVOHLC":
      case "Surface":
      case "SurfaceTopView":
      case "SurfaceTopViewWireframe":
      case "SurfaceWireframe":
        this.defaultOptions.push(
          "axes/categoryAxis/*",
          "axes/valueAxis/*",
          "axes/categoryAxis/format/*",
          "axes/valueAxis/format/*",
          "axes/categoryAxis/format/line/*",
          "axes/categoryAxis/format/font/*",
          "axes/valueAxis/format/line/*",
          "axes/valueAxis/format/font/*",
          "series/items",
        );
        break;

      case "Radar":
      case "RadarFilled":
      case "RadarMarkers":
        this.defaultOptions.push(
          "axes/valueAxis/*",
          "axes/valueAxis/format/*",
          "axes/valueAxis/format/line/*",
          "axes/valueAxis/format/font/*",
          "series/items",
        );
        break;

      case "Map":
      case "RegionMap":
        this.defaultOptions.push("series/items");
        break;

      case "BarClustered":
      case "BarStacked":
      case "BarStacked100":
      case "ConeBarClustered":
      case "ConeBarStacked":
      case "ConeBarStacked100":
      case "CylinderBarClustered":
      case "CylinderBarStacked":
      case "CylinderBarStacked100":
      case "PyramidBarClustered":
      case "PyramidBarStacked":
      case "PyramidBarStacked100":
      case "BarOfPie":
        this.defaultOptions.push(
          "axes/categoryAxis/*",
          "axes/categoryAxis/format/*",
          "axes/valueAxis/*",
          "axes/valueAxis/format/*",
          "axes/categoryAxis/format/line/*",
          "axes/categoryAxis/format/font/*",
          "axes/valueAxis/format/line/*",
          "axes/valueAxis/format/font/*",
          "series/items",
        );
        break;

      case "LineMarkers":
      case "LineMarkersStacked":
      case "LineMarkersStacked100":
      case "Pareto":
      case "PieExploded":
      case "PieOfPie":
      case "DoughnutExploded":
        this.defaultOptions.push("series/items");
        break;

      default:
        break;
    }

    return this.defaultOptions;
  }

  makeChartStyle(selectedChart, chartColors) {
    const { chartColor, legendColor, plotAreaColor } = chartColors;

    this.chartStyle = {
      chartType: selectedChart.chartType,
      font: {
        name: selectedChart.format.font.name,
        size: selectedChart.format.font.size,
        color: selectedChart.format.font.color,
        bold: selectedChart.format.font.bold,
        italic: selectedChart.format.font.italic,
        underline: selectedChart.format.font.underline,
      },
      roundedCorners: selectedChart.format.roundedCorners,
      fill: {
        color: chartColor.value,
      },
      border: {
        lineStyle: selectedChart.format.border.lineStyle,
        color: selectedChart.format.border.color,
        weight: selectedChart.format.border.weight,
      },
      plotArea: {
        fill: plotAreaColor.value,
        border: {
          lineStyle: selectedChart.plotArea.format.border.lineStyle,
          color: selectedChart.plotArea.format.border.color,
          weight: selectedChart.plotArea.format.border.weight,
        },
        position: selectedChart.plotArea.position,
        height: selectedChart.plotArea.height,
        left: selectedChart.plotArea.left,
        top: selectedChart.plotArea.top,
        width: selectedChart.plotArea.width,
        insideHeight: selectedChart.plotArea.insideHeight,
        insideLeft: selectedChart.plotArea.insideLeft,
        insideTop: selectedChart.plotArea.insideTop,
        insideWidth: selectedChart.plotArea.insideWidth,
      },
      legend: {
        fill: legendColor.value,
        font: {
          name: selectedChart.legend.format.font.name,
          size: selectedChart.legend.format.font.size,
          color: selectedChart.legend.format.font.color,
          bold: selectedChart.legend.format.font.bold,
          italic: selectedChart.legend.format.font.italic,
          underline: selectedChart.legend.format.font.underline,
        },
        border: {
          lineStyle: selectedChart.legend.format.border.lineStyle,
          color: selectedChart.legend.format.border.color,
          weight: selectedChart.legend.format.border.weight,
        },
        position: selectedChart.legend.position,
      },
    };

    if (this.loadOptions.includes("axes/categoryAxis")) {
      this.chartStyle.axes = this.chartStyle.axes || {};
      this.chartStyle.axes.categoryAxis = {
        position: selectedChart.axes.categoryAxis.position,
        format: {
          line: {
            color: selectedChart.axes.categoryAxis.format.line.color,
            style: selectedChart.axes.categoryAxis.format.line.lineStyle,
            weight: selectedChart.axes.categoryAxis.format.line.weight,
          },
          font: {
            name: selectedChart.axes.categoryAxis.format.font.name,
            size: selectedChart.axes.categoryAxis.format.font.size,
            color: selectedChart.axes.categoryAxis.format.font.color,
            bold: selectedChart.axes.categoryAxis.format.font.bold,
            italic: selectedChart.axes.categoryAxis.format.font.italic,
            underline: selectedChart.axes.categoryAxis.format.font.underline,
          },
        },
      };
    }

    if (this.loadOptions.includes("axes/valueAxis")) {
      this.chartStyle.axes = this.chartStyle.axes || {};
      this.chartStyle.axes.valueAxis = {
        position: selectedChart.axes.valueAxis.position,
        format: {
          line: {
            color: selectedChart.axes.valueAxis.format.line.color,
            style: selectedChart.axes.valueAxis.format.line.lineStyle,
            weight: selectedChart.axes.valueAxis.format.line.weight,
          },
          font: {
            name: selectedChart.axes.valueAxis.format.font.name,
            size: selectedChart.axes.valueAxis.format.font.size,
            color: selectedChart.axes.valueAxis.format.font.color,
            bold: selectedChart.axes.valueAxis.format.font.bold,
            italic: selectedChart.axes.valueAxis.format.font.italic,
            underline: selectedChart.axes.valueAxis.format.font.underline,
          },
        },
      };
    }

    return this.chartStyle;
  }
}

export default ChartInfo;
