import { popUpMessage } from "./commonFuncs";

async function saveChartStylePreset(targetPreset, styleName) {
  try {
    if (styleName === "") {
      popUpMessage("loadFail", "프리셋을 정확하게 선택해주세요.");

      return;
    }

    await Excel.run(async (context) => {
      let chartStylePresets = await OfficeRuntime.storage.getItem(targetPreset);

      if (!chartStylePresets) {
        chartStylePresets = {};
      } else {
        chartStylePresets = JSON.parse(chartStylePresets);
      }

      const selectedChart = context.workbook.getActiveChart();

      if (!selectedChart) {
        popUpMessage("loadFail", "선택된 차트를 찾을 수 없습니다.");
      }

      selectedChart.load("chartType");
      await context.sync();

      const currentChartType = selectedChart.chartType;

      const propertiesToLoad = [
        "format/fill",
        "format/font",
        "format/border",
        "format/roundedCorners",
        "plotArea/format/fill",
        "plotArea/format/border",
        "plotArea/position",
        "plotArea/height",
        "plotArea/left",
        "plotArea/top",
        "plotArea/width",
        "plotArea/insideHeight",
        "plotArea/insideLeft",
        "plotArea/insideTop",
        "plotArea/insideWidth",
        "legend/format/fill",
        "legend/format/font",
        "legend/format/border",
        "legend/position",
        "seriesNameLevel",
      ];

      switch (currentChartType) {
        case Excel.ChartType.columnClustered:
        case Excel.ChartType.columnStacked:
        case Excel.ChartType.columnStacked100:
        case Excel.ChartType.line:
        case Excel.ChartType.lineStacked:
        case Excel.ChartType.lineStacked100:
        case Excel.ChartType.area:
        case Excel.ChartType.areaStacked:
        case Excel.ChartType.areaStacked100:
        case Excel.ChartType.histogram:
        case Excel.ChartType.boxWhisker:
        case Excel.ChartType.waterfall:
        case Excel.ChartType.funnel:
        case Excel.ChartType._3DArea:
        case Excel.ChartType._3DAreaStacked:
        case Excel.ChartType._3DAreaStacked100:
        case Excel.ChartType._3DColumn:
        case Excel.ChartType._3DColumnClustered:
        case Excel.ChartType._3DColumnStacked:
        case Excel.ChartType._3DColumnStacked100:
        case Excel.ChartType._3DLine:
        case Excel.ChartType._3DBarClustered:
        case Excel.ChartType._3DBarStacked:
        case Excel.ChartType._3DBarStacked100:
          propertiesToLoad.push(
            "axes/categoryAxis/format/line",
            "axes/categoryAxis/format/font",
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/categoryAxis/position",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.pie:
        case Excel.ChartType.doughnut:
        case Excel.ChartType.treemap:
        case Excel.ChartType.sunburst:
        case Excel.ChartType._3DPie:
        case Excel.ChartType._3DPieExploded:
          propertiesToLoad.push("series");
          break;

        case Excel.ChartType.scatter:
        case Excel.ChartType.bubble:
        case Excel.ChartType.xyscatter:
        case Excel.ChartType.xyscatterLines:
        case Excel.ChartType.xyscatterLinesNoMarkers:
        case Excel.ChartType.xyscatterSmooth:
        case Excel.ChartType.xyscatterSmoothNoMarkers:
          propertiesToLoad.push(
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.stockHLC:
        case Excel.ChartType.stockOHLC:
        case Excel.ChartType.stockVHLC:
        case Excel.ChartType.stockVOHLC:
        case Excel.ChartType.surface:
        case Excel.ChartType.surfaceTopView:
        case Excel.ChartType.surfaceTopViewWireframe:
        case Excel.ChartType.surfaceWireframe:
          propertiesToLoad.push(
            "axes/categoryAxis/format/line",
            "axes/categoryAxis/format/font",
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/categoryAxis/position",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.radar:
        case Excel.ChartType.radarFilled:
        case Excel.ChartType.radarMarkers:
          propertiesToLoad.push(
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.map:
        case Excel.ChartType.regionMap:
          propertiesToLoad.push("series");
          break;

        case Excel.ChartType.barClustered:
        case Excel.ChartType.barStacked:
        case Excel.ChartType.barStacked100:
        case Excel.ChartType.coneBarClustered:
        case Excel.ChartType.coneBarStacked:
        case Excel.ChartType.coneBarStacked100:
        case Excel.ChartType.cylinderBarClustered:
        case Excel.ChartType.cylinderBarStacked:
        case Excel.ChartType.cylinderBarStacked100:
        case Excel.ChartType.pyramidBarClustered:
        case Excel.ChartType.pyramidBarStacked:
        case Excel.ChartType.pyramidBarStacked100:
        case Excel.ChartType.barOfPie:
          propertiesToLoad.push(
            "axes/categoryAxis/format/line",
            "axes/categoryAxis/format/font",
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/categoryAxis/position",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.lineMarkers:
        case Excel.ChartType.lineMarkersStacked:
        case Excel.ChartType.lineMarkersStacked100:
        case Excel.ChartType.pareto:
        case Excel.ChartType.pieExploded:
        case Excel.ChartType.pieOfPie:
        case Excel.ChartType.doughnutExploded:
          propertiesToLoad.push("series");
          break;

        default:
          popUpMessage("loadFail", "지원하지 않는 차트 유형입니다.");

          return;
      }

      selectedChart.load(propertiesToLoad);
      await context.sync();

      const chartFillColor = selectedChart.format.fill.getSolidColor();
      const legendFillColor = selectedChart.legend.format.fill.getSolidColor();
      const plotAreaFillColor =
        selectedChart.plotArea.format.fill.getSolidColor();

      await context.sync();

      const chartStyle = {
        chartType: currentChartType,
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
          color: chartFillColor,
        },
        border: {
          lineStyle: selectedChart.format.border.lineStyle,
          color: selectedChart.format.border.color,
          weight: selectedChart.format.border.weight,
        },
        plotArea: {
          fill: plotAreaFillColor,
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
          fill: legendFillColor,
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
        seriesNameLevel: selectedChart.seriesNameLevel,
      };

      if (propertiesToLoad.includes("axes/categoryAxis")) {
        chartStyle.axes = chartStyle.axes || {};
        chartStyle.axes.categoryAxis = {
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

      if (propertiesToLoad.includes("axes/valueAxis")) {
        chartStyle.axes = chartStyle.axes || {};
        chartStyle.axes.valueAxis = {
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

      if (propertiesToLoad.includes("series")) {
        selectedChart.series.load("items");
        await context.sync();

        chartStyle.series = [];

        for (let i = 0; i < selectedChart.series.items.length; i += 1) {
          const series = selectedChart.series.items[i];

          series.load(["format/fill", "format/line"]);
          await context.sync();

          chartStyle.series.push(series);
        }
      }

      chartStylePresets[styleName] = chartStyle;

      await OfficeRuntime.storage.setItem(
        targetPreset,
        JSON.stringify(chartStylePresets),
      );

      popUpMessage("saveSuccess");
    });
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error.message, error.stack);
  }
}

async function loadChartStylePreset(targetPreset, styleName) {
  if (styleName === "") {
    popUpMessage("loadFail", "프리셋을 정확하게 선택해주세요!");

    return;
  }

  try {
    await Excel.run(async (context) => {
      const currentChart = context.workbook.getActiveChart();

      if (!currentChart) {
        popUpMessage("loadFail", "선택된 차트를 찾을 수 없습니다.");

        return;
      }

      currentChart.load("chartType");
      await context.sync();

      let chartStylePresets = await OfficeRuntime.storage.getItem(targetPreset);

      if (!chartStylePresets) {
        popUpMessage("loadFail", "프리셋 목록을 불러오는데 실패했습니다.");

        return;
      }

      chartStylePresets = JSON.parse(chartStylePresets);
      const chartStyle = chartStylePresets[styleName];

      if (!chartStyle) {
        popUpMessage("loadFail", "해당 프리셋을 찾을 수 없습니다.");

        return;
      }

      if (
        chartStyle.chartType &&
        currentChart.chartType !== chartStyle.chartType
      ) {
        popUpMessage(
          "loadFail",
          "차트 타입이 다릅니다. 일부 스타일이 적용되지 않을 수 있습니다.",
        );
      }

      applyBasicChartProperties(currentChart, chartStyle);
      applyLegendProperties(currentChart, chartStyle);
      applyPlotAreaProperties(currentChart, chartStyle);
      applyAxisProperties(currentChart, chartStyle);

      if (chartStyle.series && currentChart.series) {
        context.trackedObjects.add(currentChart);

        await applySeriesProperties(currentChart, chartStyle);

        context.trackedObjects.remove(currentChart);
      }

      await context.sync();

      popUpMessage("loadSuccess", "차트 서식을 적용했습니다.");
    });
  } catch (error) {
    popUpMessage("loadFail", "차트 서식 적용에 실패하였습니다.");
  }
}

function applyBasicChartProperties(currentChart, chartStyle) {
  if (chartStyle.fill.color) {
    currentChart.format.fill.setSolidColor(chartStyle.fill.color.m_value);
  } else {
    currentChart.format.fill.clear();
  }

  if (chartStyle.border) {
    if (chartStyle.border.lineStyle !== "none") {
      if (chartStyle.border.color) {
        currentChart.format.border.color = chartStyle.border.color;
      }

      if (chartStyle.border.lineStyle) {
        currentChart.format.border.lineStyle = chartStyle.border.lineStyle;
      }

      if (chartStyle.border.weight && chartStyle.border.weight > 0) {
        currentChart.format.border.weight = chartStyle.border.weight;
      }
    } else {
      currentChart.format.border.clear();
    }
  }

  if (chartStyle.font) {
    Object.keys(chartStyle.font).forEach((key) => {
      if (chartStyle.font[key] !== undefined) {
        currentChart.format.font[key] = chartStyle.font[key];
      }
    });
  }

  if (chartStyle.roundedCorners !== undefined) {
    currentChart.format.roundedCorners = chartStyle.roundedCorners;
  }
}

function applyLegendProperties(currentChart, chartStyle) {
  if (chartStyle.legend) {
    if (chartStyle.legend.fill.color) {
      currentChart.legend.format.fill.setSolidColor(
        chartStyle.legend.fill.color.m_value,
      );
    } else {
      currentChart.legend.format.fill.clear();
    }

    if (chartStyle.legend.border) {
      if (chartStyle.border.lineStyle !== "None") {
        if (chartStyle.legend.border.color) {
          currentChart.legend.format.border.color =
            chartStyle.legend.border.color;
        }

        if (chartStyle.legend.border.lineStyle) {
          currentChart.legend.format.border.lineStyle =
            chartStyle.legend.border.lineStyle;
        }

        if (chartStyle.border.weight && chartStyle.border.weight > 0) {
          currentChart.legend.format.border.weight =
            chartStyle.legend.border.weight;
        }
      } else {
        currentChart.legend.format.border.clear();
      }
    }

    if (chartStyle.legend.font) {
      Object.keys(chartStyle.legend.font).forEach((key) => {
        if (chartStyle.legend.font[key] !== undefined) {
          currentChart.legend.format.font[key] = chartStyle.legend.font[key];
        }
      });
    }

    if (chartStyle.legend.position) {
      currentChart.legend.position = chartStyle.legend.position;
    }
  }
}

function applyPlotAreaProperties(currentChart, chartStyle) {
  if (chartStyle.plotArea) {
    if (chartStyle.plotArea.fill.color) {
      currentChart.plotArea.format.fill.setSolidColor(
        chartStyle.plotArea.fill.color.m_value,
      );
    } else {
      currentChart.plotArea.format.fill.clear();
    }

    if (chartStyle.plotArea.border) {
      if (chartStyle.plotArea.border.lineStyle !== "None") {
        if (chartStyle.plotArea.border.color) {
          currentChart.plotArea.format.border.color =
            chartStyle.plotArea.border.color;
        }

        if (chartStyle.plotArea.border.lineStyle) {
          currentChart.plotArea.format.border.lineStyle =
            chartStyle.plotArea.border.lineStyle;
        }

        if (
          chartStyle.plotArea.border.weight &&
          chartStyle.plotArea.border.weight > 0
        ) {
          currentChart.plotArea.format.border.weight =
            chartStyle.plotArea.border.weight;
        }
      } else {
        currentChart.plotArea.format.border.clear();
      }
    }

    if (chartStyle.plotArea.position === "Automatic") {
      currentChart.plotArea.position = chartStyle.plotArea.position;
    } else {
      currentChart.plotArea.height = chartStyle.plotArea.height;
      currentChart.plotArea.left = chartStyle.plotArea.left;
      currentChart.plotArea.top = chartStyle.plotArea.top;
      currentChart.plotArea.width = chartStyle.plotArea.width;
      currentChart.plotArea.insideHeight = chartStyle.plotArea.insideHeight;
      currentChart.plotArea.insideLeft = chartStyle.plotArea.insideLeft;
      currentChart.plotArea.insideTop = chartStyle.plotArea.insideTop;
      currentChart.plotArea.insideWidth = chartStyle.plotArea.insideWidth;
    }
  }
}

function applyAxisProperties(currentChart, chartStyle) {
  if (chartStyle.axes) {
    if (chartStyle.axes.categoryAxis && currentChart.axes.categoryAxis) {
      applySingleAxisProperties(
        currentChart.axes.categoryAxis,
        chartStyle.axes.categoryAxis,
      );
    }

    if (chartStyle.axes.valueAxis && currentChart.axes.valueAxis) {
      applySingleAxisProperties(
        currentChart.axes.valueAxis,
        chartStyle.axes.valueAxis,
      );
    }
  }
}

function applySingleAxisProperties(axis, axisStyle) {
  if (axisStyle.format) {
    if (axisStyle.format.line.lineStyle !== "None") {
      if (axisStyle.format.line) {
        axis.format.line.color = axisStyle.format.line.color;
        axis.format.line.lineStyle = axisStyle.format.line.style;

        if (axisStyle.format.line.weight > 0) {
          axis.format.line.weight = axisStyle.format.line.weight;
        }
      }
    } else {
      axis.format.line.clear();
    }

    if (axisStyle.format.font) {
      Object.keys(axisStyle.format.font).forEach((key) => {
        if (axisStyle.format.font[key] !== undefined)
          axis.format.font[key] = axisStyle.format.font[key];
      });
    }
  }

  if (axisStyle.position) {
    axis.position = axisStyle.position;
  }
}

async function applySeriesProperties(currentChart, chartStyle) {
  if (chartStyle.series && currentChart.series) {
    await currentChart.series.load("items");
    await currentChart.context.sync();

    const seriesArray = Array.isArray(chartStyle.series)
      ? chartStyle.series
      : Object.values(chartStyle.series);

    for (
      let index = 0;
      index < Math.min(seriesArray.length, currentChart.series.items.length);
      index += 1
    ) {
      const series = currentChart.series.items[index];

      if (series) {
        series.load(["format/fill", "format/line"]);
      }
    }

    await currentChart.context.sync();

    for (
      let index = 0;
      index < Math.min(seriesArray.length, currentChart.series.items.length);
      index += 1
    ) {
      const seriesStyle = seriesArray[index];
      const series = currentChart.series.items[index];

      if (seriesStyle.format) {
        if (seriesStyle.format.fill) {
          if (seriesStyle.format.fill.color) {
            series.format.fill.setSolidColor(seriesStyle.format.fill.color);
          } else {
            series.format.fill.clear();
          }
        }

        if (seriesStyle.format.line.style !== "None") {
          if (seriesStyle.format.line.color) {
            series.format.line.color = seriesStyle.format.line.color;
          }

          if (seriesStyle.format.line.style) {
            series.format.line.lineStyle = seriesStyle.format.line.style;
          }

          if (
            seriesStyle.format.line.weight &&
            seriesStyle.format.line.weight > 0
          ) {
            series.format.line.weight = seriesStyle.format.line.weight;
          }
        } else {
          series.format.line.clear();
        }
      }
    }
  }
}

export { saveChartStylePreset, loadChartStylePreset };
