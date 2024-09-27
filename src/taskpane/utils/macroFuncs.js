import { storeCellStyle, applyCellStyle } from "./cellStyleFuncs";
import { updateState } from "./commonFuncs";

let worksheetChangedHandler;
let tableChangedHandler;
let chartAddedHandler;
let tableAddedHandler;
let formatChangedHandler;
let actions = [];

async function manageRecording(isRecording, presetName) {
  if (presetName === "") {
    const warningMessage = {
      type: "warning",
      title: "접근 오류: ",
      body: `프리셋을 선택해주세요!`,
    };

    updateState("setMessageList", warningMessage);
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const newTables = context.workbook.tables;

    if (isRecording) {
      actions = [];
      tableChangedHandler = newTables.onChanged.add((event) =>
        onWorksheetChanged(event, presetName),
      );
      chartAddedHandler = sheet.charts.onAdded.add((event) =>
        onWorksheetChanged(event, presetName),
      );
      tableAddedHandler = sheet.tables.onAdded.add((event) =>
        onWorksheetChanged(event, presetName),
      );
      formatChangedHandler = sheet.onFormatChanged.add(async (event) =>
        onWorksheetChanged(event, presetName),
      );
      worksheetChangedHandler = context.workbook.worksheets.onChanged.add(
        (event) => onWorksheetChanged(event, presetName),
      );

      let allMacroPresets =
        await OfficeRuntime.storage.getItem("allMacroPresets");
      allMacroPresets = allMacroPresets ? JSON.parse(allMacroPresets) : {};
      allMacroPresets[presetName] = { actions: [] };

      await OfficeRuntime.storage.setItem(
        "allMacroPresets",
        JSON.stringify(allMacroPresets),
      );
      await context.sync();
    } else {
      if (worksheetChangedHandler) {
        await Excel.run(worksheetChangedHandler.context, async (ctx) => {
          worksheetChangedHandler.remove();
          await ctx.sync();
        });
      }

      if (tableChangedHandler) {
        await Excel.run(tableChangedHandler.context, async (ctx) => {
          tableChangedHandler.remove();
          await ctx.sync();
        });
      }

      if (chartAddedHandler) {
        await Excel.run(chartAddedHandler.context, async (ctx) => {
          chartAddedHandler.remove();
          await ctx.sync();
        });
      }

      if (tableAddedHandler) {
        await Excel.run(tableAddedHandler.context, async (ctx) => {
          tableAddedHandler.remove();
          await ctx.sync();
        });
      }

      if (formatChangedHandler) {
        await Excel.run(formatChangedHandler.context, async (ctx) => {
          formatChangedHandler.remove();
          await ctx.sync();
        });
      }

      worksheetChangedHandler = null;
      tableChangedHandler = null;
      chartAddedHandler = null;
      tableAddedHandler = null;
      formatChangedHandler = null;

      await saveMacro(presetName);
    }
  }).catch((error) => {
    const warningMessage = {
      type: "warning",
      title: "에러 발생",
      body: `녹화를 시작할 수 없습니다. ${error.message}`,
    };

    updateState("setMessageList", warningMessage);
  });
}

async function onWorksheetChanged(event, presetName) {
  const action = { type: event.type };
  let cellStyleData = null;
  let allMacroPresets = await OfficeRuntime.storage.getItem("allMacroPresets");
  let warningMessage = {};
  allMacroPresets = allMacroPresets ? JSON.parse(allMacroPresets) : {};

  if (!allMacroPresets[presetName].actions) {
    allMacroPresets[presetName] = { actions: [], cellStyles: {} };
  }

  try {
    switch (event.type) {
      case "WorksheetChanged":
        action.address = event.address;
        action.details = {
          value: event.details.valueAfter ? event.details.valueAfter : "",
        };
        break;

      case "WorksheetFormatChanged":
        action.address = event.address;
        cellStyleData = await storeCellStyle(
          event.address,
          "allMacroPresets",
          true,
        );
        action.cellStyle = cellStyleData;
        break;

      case "TableChanged":
        action.tableId = event.tableId;
        action.changeType = event.changeType;
        action.address = event.address;
        action.details = event.details;
        break;

      case "ChartAdded":
        action.chartId = event.chartId;
        await onChartAdded(action);
        break;

      case "TableAdded":
        action.tableId = event.tableId;
        [action.address, action.showHeaders] = await onTableAdded(action);
        break;

      default:
        warningMessage = {
          type: "warning",
          title: "접근 오류",
          body: "지원하지 않는 형식입니다.",
        };

        updateState("setMessageList", warningMessage);
        break;
    }
  } catch (e) {
    warningMessage = {
      type: "error",
      title: "에러 발생: ",
      body: `기록 중 예상치 못한 에러가 발생했습니다. ${e.message}`,
    };

    updateState("setMessageList", warningMessage);
  }

  if (action.chartType === "Unknown") {
    warningMessage = {
      type: "warning",
      title: "지원하지 않는 차트: ",
      body: "매크로 설정에서 직접 차트 타입을 입력해주세요.",
    };

    updateState("setMessageList", warningMessage);
  }

  actions.push(action);
}

async function saveMacro(presetName) {
  try {
    let allMacroPresets =
      await OfficeRuntime.storage.getItem("allMacroPresets");
    allMacroPresets = allMacroPresets ? JSON.parse(allMacroPresets) : {};

    if (!allMacroPresets[presetName].actions) {
      allMacroPresets[presetName] = { actions: [], cellStyles: {} };
    }

    allMacroPresets[presetName].actions = actions;

    await OfficeRuntime.storage.setItem(
      "allMacroPresets",
      JSON.stringify(allMacroPresets),
    );

    const successMessage = {
      type: "success",
      title: "저장 완료: ",
      body: "매크로를 기록했습니다.",
    };

    updateState("setMessageList", successMessage);
  } catch (error) {
    const warningMessage = {
      type: "warning",
      title: "저장 실패",
      body: `매크로 기록에 실패했습니다. ${error.message}`,
    };

    updateState("setMessageList", warningMessage);
  }
}

async function onChartAdded(action) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.getItem(action.chartId);
      const dataRange = [];

      chart.load([
        "top",
        "left",
        "height",
        "width",
        "series/items",
        "chartType",
      ]);

      await context.sync();

      if (chart.chartType) {
        for (let i = 0; i < chart.series.count; i += 1) {
          const series = chart.series.getItemAt(i);
          let valuesDataSource;

          try {
            valuesDataSource = series.getDimensionDataSourceString("Values");
          } catch (error) {
            try {
              series.load("values");
              await context.sync();

              valuesDataSource = { value: series.values.address };
            } catch (innerError) {
              valuesDataSource = { value: "Unknown" };
            }
          }

          await context.sync();

          dataRange.push({
            address: valuesDataSource.value.split("!")[1],
          });
        }
      } else {
        const warningMessage = {
          type: "warning",
          title: "형식 오류: ",
          body: "지원하지 않는 차트 형식입니다. 매크로 설정에서 반드시 차트 형식을 지정해주세요!",
        };

        updateState("setMessageList", warningMessage);
      }

      action.chartType = chart.chartType || "Unknown";
      action.position = { top: chart.top, left: chart.left };
      action.size = { height: chart.height, width: chart.width };
      action.dataRange = dataRange.map((range) => range.address);

      return action;
    });
  } catch (e) {
    const warningMessage = {
      type: "warning",
      title: "저장 실패",
      body: `매크로 기록에 실패했습니다. ${e.message}`,
    };

    updateState("setMessageList", warningMessage);
  }
}

async function onTableAdded(action) {
  try {
    let tableAttributes = [];

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem(action.tableId);
      const tableRange = table.getRange();

      tableRange.load("address");
      table.load("showHeaders");
      await context.sync();

      const { showHeaders } = table;
      tableAttributes = [tableRange.address.split("!")[1], showHeaders];
    });

    return tableAttributes;
  } catch (e) {
    const warningMessage = {
      type: "warning",
      title: "저장 실패",
      body: `매크로 기록에 실패했습니다. ${e.message}`,
    };

    updateState("setMessageList", warningMessage);

    return null;
  }
}

async function macroPlay(presetName) {
  try {
    await Excel.run(async (context) => {
      const allMacroPresets =
        await OfficeRuntime.storage.getItem("allMacroPresets");
      let warningMessage = {};

      if (!allMacroPresets) {
        throw new Error("No macros found.");
      }

      const parsedPresets = JSON.parse(allMacroPresets);
      const presetData = parsedPresets[presetName];

      if (!presetData || !presetData.actions) {
        throw new Error(`No actions found for preset: ${presetName}`);
      }

      for (const action of presetData.actions) {
        switch (action.type) {
          case "WorksheetChanged":
            await applyWorksheetChange(context, action);
            break;

          case "WorksheetFormatChanged":
            await applyCellStyle(
              action.address,
              "allMacroPresets",
              false,
              action.cellStyle,
            );
            break;

          case "TableChanged":
            await applyTableChange(context, action);
            break;

          case "ChartAdded":
            await applyChartAdded(context, action);
            break;

          case "TableAdded":
            await applyTableAdded(context, action);
            break;

          default:
            warningMessage = {
              type: "warning",
              title: "형식 오류",
              body: `지원하지 않는 형식의 기록입니다. 매크로 재생에 실패했습니다. `,
            };

            updateState("setMessageList", warningMessage);
            break;
        }
      }

      await context.sync();
    });
  } catch (error) {
    const warningMessage = {
      type: "warning",
      title: "재생 실패",
      body: "매크로 재생에 실패했습니다. 차트 형식, 혹은 지원되는 기록인지 확인해주세요.",
    };

    updateState("setMessageList", warningMessage);
  }
}

async function applyWorksheetChange(context, action) {
  if (action.details && action.details.value) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(action.address);

    range.values = [[action.details.value]];
    await context.sync();
  }
}

async function applyTableChange(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let warningMessage = {};

  switch (action.changeType) {
    case "RangeEdited":
      if (action.details && action.details.valueAfter) {
        const range = sheet.getRange(action.address);
        range.values = [[action.details.valueAfter]];

        await context.sync();
      }
      break;

    default:
      warningMessage = {
        type: "warning",
        title: "접근 오류",
        body: "지원되지 않는 표 이벤트 입니다.",
      };

      updateState("setMessageList", warningMessage);

      break;
  }
}

async function applyChartAdded(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let mergedRange = null;

  if (action.dataRange[0].includes(":")) {
    mergedRange = `${action.dataRange[0].split(":")[0]}:${action.dataRange[action.dataRange.length - 1].split(":")[1]}`;
  } else {
    mergedRange = `${action.dataRange[0]}:${action.dataRange[action.dataRange.length - 1]}`;
  }

  const chart = sheet.charts.add(action.chartType, sheet.getRange(mergedRange));

  chart.top = action.position.top;
  chart.left = action.position.left;
  chart.height = action.size.height;
  chart.width = action.size.width;

  await context.sync();
}

async function applyTableAdded(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.tables.add(action.address);
  await context.sync();
}

export { manageRecording, macroPlay, saveMacro };
