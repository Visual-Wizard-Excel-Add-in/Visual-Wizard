import { storeCellStyle, applyCellStyle } from "./cellStyleFuncs";
import {
  getSelectRangeValue,
  popUpMessage,
  removeHandler,
} from "./commonFuncs";
import useHandlerStore from "../store/handlerStore";

const actions = [];

async function manageRecording(isRecording, presetName) {
  if (presetName === "") {
    popUpMessage("loadFail", "프리셋을 정확하게 선택해주세요!");
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const { tables } = context.workbook;
    const handlers = {
      tableChangedHandler: {
        target: tables,
        eventName: "onChanged",
        setter: "setTableChangedHandler",
      },
      chartAddedHandler: {
        target: sheet.charts,
        eventName: "onAdded",
        setter: "setTableAddedHandler",
      },
      tableAddedHandler: {
        target: sheet.tables,
        eventName: "onAdded",
        setter: "setChartAddedHandler",
      },
      formatChangedHandler: {
        target: sheet,
        eventName: "onFormatChanged",
        setter: "setFormatChangedHandler",
      },
      worksheetChangedHandler: {
        target: sheet,
        eventName: "onChanged",
        setter: "setWorksheetChangedHandler",
      },
    };

    if (isRecording) {
      addHandler();

      const allMacroPresets = await getStorage("allMacroPresets");
      allMacroPresets[presetName] = { actions: [] };

      await setStorage("allMacroPresets", allMacroPresets);

      await await context.sync();
    } else {
      await removeEventHandler();

      await saveMacro(presetName);
    }

    function addHandler() {
      Object.keys(handlers).forEach((handler) => {
        const eventHandler = handlers[handler];

        useHandlerStore
          .getState()
          [
            eventHandler.setter
          ](eventHandler.target[eventHandler.eventName].add((event) => onWorksheetChanged(event, presetName)));
      });
    }

    async function removeEventHandler() {
      const requests = Object.keys(handlers).map((handler) => {
        return removeHandler(
          useHandlerStore.getState()[handler],
          handlers[handler].setter,
        );
      });

      await Promise.allSettled(requests);
    }

    async function getStorage(data) {
      const storageData = await OfficeRuntime.storage.getItem(data);

      return storageData ? JSON.parse(storageData) : {};
    }

    async function setStorage(key, data) {
      await OfficeRuntime.storage.setItem(key, JSON.stringify(data));
    }
  }).catch((error) => {
    popUpMessage("workFail", `녹화를 시작할 수 없습니다. ${error.message}`);

    throw new Error(error.message);
  });
}

async function onWorksheetChanged(event, presetName) {
  let allMacroPresets = await OfficeRuntime.storage.getItem("allMacroPresets");

  allMacroPresets = allMacroPresets ? JSON.parse(allMacroPresets) : {};

  if (!allMacroPresets[presetName].actions) {
    allMacroPresets[presetName] = { actions: [], cellStyles: {} };
  }

  const action = { type: event.type };

  try {
    await recordAction();
  } catch (error) {
    popUpMessage(
      "workFail",
      `기록 중 예상치 못한 에러가 발생했습니다. ${error.message}`,
    );
  }

  if (action.chartType === "Unknown") {
    popUpMessage("loadFail", "매크로 설정에서 차트 타입을 변경해주세요.");
  }

  actions.push(action);

  async function recordAction() {
    switch (event.type) {
      case "WorksheetChanged":
        action.address = event.address;
        action.details = {
          value: event.details
            ? event.details.valueAfter
            : await getSelectRangeValue(),
        };
        break;

      case "WorksheetFormatChanged":
        action.address = event.address;
        action.cellStyle = await storeCellStyle(
          event.address,
          "allMacroPresets",
          true,
        );
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
        popUpMessage("loadFail", "지원하지 않는 형식입니다.");
        break;
    }
  }
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

    popUpMessage("saveSuccess", "매크로를 기록했습니다.");
  } catch (error) {
    popUpMessage("saveFail", `매크로 기록에 실패했습니다. ${error.message}`);
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
        popUpMessage("workFail", "매크로 설정에서 차트 타입을 변경해주세요.");
      }

      action.chartType = chart.chartType || "Unknown";
      action.position = { top: chart.top, left: chart.left };
      action.size = { height: chart.height, width: chart.width };
      action.dataRange = dataRange.map((range) => range.address);

      return action;
    });
  } catch (error) {
    popUpMessage("saveFail", `매크로 기록에 실패했습니다. ${error.message}`);

    throw new Error(error.message);
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
  } catch (error) {
    popUpMessage("saveFail", `매크로 기록에 실패했습니다. ${error.message}`);

    throw new Error(error.message);
  }
}

async function macroPlay(presetName) {
  try {
    await Excel.run(async (context) => {
      const allMacroPresets =
        await OfficeRuntime.storage.getItem("allMacroPresets");

      if (!allMacroPresets) {
        throw new Error("No macros found.");
      }

      const parsedPresets = JSON.parse(allMacroPresets);
      const presetData = parsedPresets[presetName];

      if (!presetData || !presetData.actions) {
        throw new Error(`No actions found for preset: ${presetName}`);
      }

      for (const action of presetData.actions) {
        await replayRecords(action);
      }

      await context.sync();

      async function replayRecords(action) {
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
            popUpMessage("workFail", "지원하지 않는 형식의 기록입니다.");
            break;
        }
      }
    });
  } catch (error) {
    popUpMessage("workFail", "지원하는 타입의 기록인지 확인해주세요.");
  }
}

async function applyWorksheetChange(context, action) {
  if (action.details && action.details.value) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(action.address);

    if (typeof action.details.value === "object") {
      range.values = action.details.value;
    } else {
      range.values = [[action.details.value]];
    }
    await context.sync();
  }
}

async function applyTableChange(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  switch (action.changeType) {
    case "RangeEdited":
      if (action.details && action.details.valueAfter) {
        const range = sheet.getRange(action.address);
        range.values = [[action.details.valueAfter]];

        await context.sync();
      }
      break;

    default:
      popUpMessage("loadFail", "지원하지 않는 표 이벤트입니다.");
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
