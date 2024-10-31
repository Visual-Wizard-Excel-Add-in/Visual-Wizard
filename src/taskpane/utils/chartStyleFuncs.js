import { popUpMessage } from "./commonFuncs";
import ChartInfo from "../classes/ChartInfo";

async function copyChartStyle(targetPreset, styleName) {
  if (styleName === "") {
    popUpMessage("loadFail", "프리셋을 정확하게 선택해주세요.");

    return;
  }

  try {
    await Excel.run(async (context) => {
      const chartStylePresets = await loadStorage(targetPreset);
      const selectedChart = context.workbook.getActiveChart();

      if (!selectedChart) {
        popUpMessage("loadFail", "선택된 차트를 찾을 수 없습니다.");
      }

      selectedChart.load(["*", "chartType"]);
      await context.sync();

      const chart = new ChartInfo(selectedChart.chartType);
      const chartColors = getChartColors();

      selectedChart.load(chart.loadOptions);
      await context.sync();

      chartStylePresets[styleName] = chart.makeChartStyle(
        selectedChart,
        chartColors,
      );

      if (chart.loadOptions.includes("series/items")) {
        await makeSeriesStyles();
      }

      await OfficeRuntime.storage.setItem(
        targetPreset,
        JSON.stringify(chartStylePresets),
      );

      popUpMessage("saveSuccess");

      function getChartColors() {
        return {
          chartColor: selectedChart.format.fill.getSolidColor(),
          legendColor: selectedChart.legend.format.fill.getSolidColor(),
          plotAreaColor: selectedChart.plotArea.format.fill.getSolidColor(),
        };
      }

      async function makeSeriesStyles() {
        chart.chartStyle.series = [];

        for (let i = 0; i < selectedChart.series.items.length; i += 1) {
          try {
            const series = selectedChart.series.items[i];
            const color = series.format.fill.getSolidColor();

            series.load(["format/line", "*"]);
            await context.sync();

            chart.chartStyle.series.push({
              line: series.format.line,
              fill: color.value,
            });
          } catch (error) {
            popUpMessage("default", "단색 속성만 저장 가능합니다");

            const series = selectedChart.series.items[i];

            series.load(["format/line", "*"]);
            await context.sync();

            chart.chartStyle.series.push({ line: series.format.line });
          }
        }
      }
    });
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error);
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

      const chartStylePresets = await loadStorage();

      if (!chartStylePresets) {
        popUpMessage("loadFail", "프리셋 목록을 불러오는데 실패했습니다.");

        return;
      }

      const savedStyle = chartStylePresets[styleName];

      if (!savedStyle) {
        popUpMessage("loadFail", "해당 프리셋을 찾을 수 없습니다.");

        return;
      }

      if (currentChart.chartType !== savedStyle.chartType) {
        popUpMessage(
          "loadFail",
          "차트 타입이 다릅니다. 일부 스타일이 적용되지 않을 수 있습니다.",
        );
      }

      const applyFuncs = [
        applyBasicChartProperties,
        applyLegendProperties,
        applyPlotAreaProperties,
        applyAxisProperties,
      ];

      applyFuncs.forEach((func) => func(currentChart, savedStyle));

      if (savedStyle.series) {
        await applySeriesProperties(currentChart, savedStyle);
      }

      await context.sync();

      popUpMessage("loadSuccess", "차트 서식을 적용했습니다.");
    });
  } catch (error) {
    popUpMessage("workFail", "차트 서식 적용에 실패하였습니다.");

    throw new Error(error);
  }

  async function loadStorage() {
    const loadedPresets = await OfficeRuntime.storage.getItem(targetPreset);

    if (!loadedPresets) {
      return null;
    }

    return JSON.parse(loadedPresets);
  }
}

function applyBasicChartProperties(target, savedStyle) {
  if (savedStyle.fill.color) {
    target.format.fill.setSolidColor(savedStyle.fill.color);
  } else {
    target.format.fill.clear();
  }

  if (savedStyle.border) {
    if (savedStyle.border.lineStyle !== "None") {
      const { color, lineStyle, weight } = savedStyle.border;

      if (color) {
        target.format.border.color = color;
      }

      if (lineStyle) {
        target.format.border.lineStyle = lineStyle;
      }

      if (weight > 0) {
        target.format.border.weight = weight;
      }
    } else {
      target.format.border.clear();
    }
  }

  if (savedStyle.font) {
    Object.keys(savedStyle.font).forEach((key) => {
      if (savedStyle.font[key]) {
        target.format.font[key] = savedStyle.font[key];
      }
    });
  }

  if (savedStyle.roundedCorners) {
    target.format.roundedCorners = savedStyle.roundedCorners;
  }
}

function applyLegendProperties(target, savedStyle) {
  if (savedStyle.legend) {
    if (savedStyle.legend.fill) {
      target.legend.format.fill.setSolidColor(savedStyle.legend.fill);
    } else {
      target.legend.format.fill.clear();
    }

    if (savedStyle.legend.border) {
      if (savedStyle.border.lineStyle !== "None") {
        const { color, lineStyle, weight } = savedStyle.legend.border;

        if (color) {
          target.legend.format.border.color = color;
        }

        if (lineStyle) {
          target.legend.format.border.lineStyle = lineStyle;
        }

        if (weight > 0) {
          target.legend.format.border.weight = weight;
        }
      } else {
        target.legend.format.border.clear();
      }
    }

    if (savedStyle.legend.font) {
      Object.keys(savedStyle.legend.font).forEach((key) => {
        if (savedStyle.legend.font[key]) {
          target.legend.format.font[key] = savedStyle.legend.font[key];
        }
      });
    }

    if (savedStyle.legend.position) {
      target.legend.position = savedStyle.legend.position;
    }
  }
}

function applyPlotAreaProperties(target, savedStyle) {
  if (savedStyle.plotArea) {
    if (savedStyle.plotArea.fill) {
      target.plotArea.format.fill.setSolidColor(savedStyle.plotArea.fill);
    } else {
      target.plotArea.format.fill.clear();
    }

    if (savedStyle.plotArea.border) {
      if (savedStyle.plotArea.border.lineStyle !== "None") {
        const { color, lineStyle, weight } = savedStyle.plotArea.border;
        if (color) {
          target.plotArea.format.border.color = color;
        }

        if (lineStyle) {
          target.plotArea.format.border.lineStyle = lineStyle;
        }

        if (weight > 0) {
          target.plotArea.format.border.weight = weight;
        }
      } else {
        target.plotArea.format.border.clear();
      }
    }

    if (savedStyle.plotArea.position === "Automatic") {
      target.plotArea.position = savedStyle.plotArea.position;
    } else {
      const positions = [
        "height",
        "left",
        "top",
        "width",
        "insideHeight",
        "insideLeft",
        "insideTop",
        "insideWidth",
      ];

      const sourceStyle = Object.fromEntries(
        positions.map((position) => [position, savedStyle.plotArea[position]]),
      );

      Object.assign(target.plotArea, sourceStyle);
    }
  }
}

function applyAxisProperties(target, savedStyle) {
  if (savedStyle.axes) {
    const { categoryAxis: targetCategory, valueAxis: targetValue } =
      target.axes;
    const { categoryAxis: savedCategory, valueAxis: savedValue } =
      savedStyle.axes;

    if (savedCategory && targetCategory) {
      applySingleAxisProperties(targetCategory, savedCategory);
    }

    if (savedValue && targetValue) {
      applySingleAxisProperties(targetValue, savedValue);
    }
  }
}

function applySingleAxisProperties(axis, axisStyle) {
  if (axisStyle.format) {
    if (axisStyle.format.line.lineStyle !== "None") {
      const { color, style, weight } = axisStyle.format.line;

      if (color) {
        axis.format.line.color = color;
      }

      if (style) {
        axis.format.line.lineStyle = style;
      }

      if (weight > 0) {
        axis.format.line.weight = weight;
      }
    } else {
      axis.format.line.clear();
    }

    if (axisStyle.format.font) {
      Object.keys(axisStyle.format.font).forEach((key) => {
        if (axisStyle.format.font[key])
          axis.format.font[key] = axisStyle.format.font[key];
      });
    }
  }

  if (axisStyle.position) {
    axis.position = axisStyle.position;
  }
}

async function applySeriesProperties(target, savedStyle) {
  if (savedStyle.series && target.series) {
    target.series.load("items");
    await target.context.sync();

    for (
      let i = 0;
      i < Math.min(savedStyle.series.length, target.series.items.length);
      i += 1
    ) {
      const seriesStyle = savedStyle.series[i];
      const series = target.series.items[i];

      if (seriesStyle) {
        if (seriesStyle.fill) {
          series.format.fill.setSolidColor(seriesStyle.fill);
        }

        if (seriesStyle.lineStyle !== "None") {
          const { color, lineStyle, weight } = seriesStyle.line;

          if (color) {
            series.format.line.color = color;
          }

          if (lineStyle) {
            series.format.line.lineStyle = lineStyle;
          }

          if (weight > 0) {
            series.format.line.weight = weight;
          }
        } else {
          series.format.line.clear();
        }
      }
    }
  }
}

export { copyChartStyle, pasteChartStyle };
