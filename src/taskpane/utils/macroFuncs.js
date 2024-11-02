import { restoreCellStyle } from "./cellStyleFuncs";
import useTotalStore from "../store/useTotalStore";
import MacroAction from "../classes/MacroActions";

const actions = [];

async function manageRecording(isRecording, presetName) {
  if (presetName === "") {
    popUpMessage("loadFail", "프리셋을 정확하게 선택해주세요!");
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const { tables } = context.workbook;
      const MACRO_HANDLERS = {
        tableChangedHandler: {
          target: tables.onChanged,
          setter: "setTableChangedHandler",
        },
        chartAddedHandler: {
          target: sheet.charts.onAdded,
          setter: "setTableAddedHandler",
        },
        tableAddedHandler: {
          target: sheet.tables.onAdded,
          setter: "setChartAddedHandler",
        },
        formatChangedHandler: {
          target: sheet.onFormatChanged,
          setter: "setFormatChangedHandler",
        },
        worksheetChangedHandler: {
          target: sheet.onChanged,
          setter: "setWorksheetChangedHandler",
        },
      };

      const allMacroPresets = await getStorage("allMacroPresets");

      if (isRecording) {
        addHandler(MACRO_HANDLERS);

        allMacroPresets[presetName] = { actions: [], cellStyles: {} };

        await setStorage(allMacroPresets);
      } else {
        await removeEventHandler(MACRO_HANDLERS);

        allMacroPresets[presetName].actions = actions;

        await setStorage(allMacroPresets);
        popUpMessage("saveSuccess", "매크로를 기록했습니다.");
      }

      await context.sync();
    });
  } catch (error) {
    popUpMessage("workFail", `녹화를 시작할 수 없습니다. ${error.message}`);

    throw new Error(error.message);
  }

  function addHandler(MACRO_HANDLERS) {
    Object.keys(MACRO_HANDLERS).forEach((handler) => {
      const eventHandler = MACRO_HANDLERS[handler];

      useTotalStore
        .getState()
        [
          eventHandler.setter
        ](eventHandler.target.add((event) => onWorksheetChanged(event)));
    });
  }

  async function removeEventHandler(MACRO_HANDLERS) {
    const requests = Object.keys(MACRO_HANDLERS).map((handler) => {
      return removeHandler(
        useTotalStore.getState()[handler],
        MACRO_HANDLERS[handler].setter,
      );
    });

    await Promise.allSettled(requests);
  }

  async function getStorage() {
    return (
      JSON.parse(await OfficeRuntime.storage.getItem("allMacroPresets")) || {}
    );
  }

  async function setStorage(data) {
    await OfficeRuntime.storage.setItem(
      "allMacroPresets",
      JSON.stringify(data),
    );
  }
}

async function onWorksheetChanged(event) {
  try {
    const action = new MacroAction(event);

    if (action.chartType === "Unknown") {
      popUpMessage("loadFail", "매크로 설정에서 차트 타입을 변경해주세요.");
    }

    actions.push(action);
  } catch (error) {
    popUpMessage(
      "workFail",
      `기록 중 예상치 못한 에러가 발생했습니다. ${error.message}`,
    );
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
            await restoreCellStyle(
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

export { manageRecording, macroPlay };
