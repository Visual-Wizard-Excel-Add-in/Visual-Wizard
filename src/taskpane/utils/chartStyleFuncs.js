import { popUpMessage } from "./commonFuncs";
import ChartInfo from "../classes/ChartInfo";

async function copyChartStyle(targetPreset, styleName) {
  try {
    if (styleName === "") {
      popUpMessage("loadFail", "프리셋을 정확하게 선택해주세요.");

      return;
    }

    await Excel.run(async (context) => {
      const chartStylePresets = await loadStorage(targetPreset);
      const selectedChart = context.workbook.getActiveChart();

      if (!selectedChart) {
        popUpMessage("loadFail", "선택된 차트를 찾을 수 없습니다.");
      }

      selectedChart.load(["*", "chartType"]);
      await context.sync();

      const chart = new ChartInfo(selectedChart.chartType);

      selectedChart.load(chart.loadOptions);

      const chartFillColor = selectedChart.format.fill.getSolidColor();
      const legendFillColor = selectedChart.legend.format.fill.getSolidColor();
      const plotAreaFillColor =
        selectedChart.plotArea.format.fill.getSolidColor();

      await context.sync();

      if (chart.loadOptions.includes("series")) {
        await makeSeriesStyles();
      }

      chartStylePresets[styleName] = chart.makeChartStyle(
        selectedChart,
        chartFillColor,
        legendFillColor,
        plotAreaFillColor,
      );

      await OfficeRuntime.storage.setItem(
        targetPreset,
        JSON.stringify(chartStylePresets),
      );

      popUpMessage("saveSuccess");

      async function makeSeriesStyles() {
        selectedChart.series.load("items");
        await context.sync();

        chart.chartStyle.series = [];

        for (let i = 0; i < selectedChart.series.items.length; i += 1) {
          const series = selectedChart.series.items[i];

          series.load(["format/fill", "format/line"]);
          await context.sync();

          chart.chartStyle.series.push(series);
        }
      }
    });
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error.message, error.stack);
  }

  async function loadStorage() {
    const loadedPresets = await OfficeRuntime.storage.getItem(targetPreset);

    if (!loadedPresets) {
      return {};
    }

    return JSON.parse(loadedPresets);
  }
}

async function pasteChartStyle(targetPreset, styleName) {
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
        await applySeriesProperties(currentChart, chartStyle);
      }

      await context.sync();

      popUpMessage("loadSuccess", "차트 서식을 적용했습니다.");
    });
  } catch (error) {
    popUpMessage("workFail", "차트 서식 적용에 실패하였습니다.");
  }
}

function applyBasicChartProperties(currentChart, chartStyle) {
  if (chartStyle.fill.color) {
    currentChart.format.fill.setSolidColor(chartStyle.fill.color);
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
    if (chartStyle.legend.fill) {
      currentChart.legend.format.fill.setSolidColor(chartStyle.legend.fill);
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
    if (chartStyle.plotArea.fill) {
      currentChart.plotArea.format.fill.setSolidColor(chartStyle.plotArea.fill);
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

export { copyChartStyle, pasteChartStyle };
