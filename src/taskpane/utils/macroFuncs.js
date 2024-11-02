import { restoreCellStyle } from "./cellStyleFuncs";
import { updateState, popUpMessage, removeHandler } from "./commonFuncs";
import useTotalStore from "../store/useTotalStore";
import MacroAction from "../classes/MacroAction";

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
      const { setter, target } = MACRO_HANDLERS[handler];

      updateState(
        setter,
        target.add((event) => onWorksheetChanged(event)),
      );
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
      const savedPresets = JSON.parse(
        await OfficeRuntime.storage.getItem("allMacroPresets"),
      );

      if (!savedPresets) {
        throw new Error("No macros found.");
      }

      const presetData = savedPresets[presetName];

      if (!presetData || !presetData.actions) {
        popUpMessage("loadFail", "녹회된 내용이 없습니다");

        return;
      }

      for (const action of presetData.actions) {
        await replayRecords(action, context);
      }

      await context.sync();
    });
  } catch (error) {
    popUpMessage("workFail", "지원하는 타입의 기록인지 확인해주세요.");
  }

  async function replayRecords(action, context) {
    const applyFuncs = {
      WorksheetChanged: () => applyWorksheetChange(context, action),
      WorksheetFormatChanged: () =>
        restoreCellStyle(
          action.address,
          "allMacroPresets",
          false,
          action.cellStyle,
        ),
      TableChanged: () => applyTableChange(context, action),
      ChartAdded: () => applyChartAdded(context, action),
      TableAdded: () => applyTableAdded(context, action),
    };

    if (applyFuncs[action.type]) {
      await applyFuncs[action.type]();
    } else {
      popUpMessage("workFail", "지원하지 않는 형식의 기록입니다.");
    }
  }
}

async function applyWorksheetChange(context, action) {
  if (action.details && action.details.value) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(action.address);
    const isRangeValue = typeof action.details.value === "object";

    if (isRangeValue) {
      range.values = action.details.value;
    } else {
      range.values = [[action.details.value]];
    }

    await context.sync();
  }
}

async function applyTableChange(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const isEditValue =
    action.changeType === "RangeEdited" &&
    action.details &&
    action.details.valueAfter;

  if (isEditValue) {
    const range = sheet.getRange(action.address);
    range.values = [[action.details.valueAfter]];

    await context.sync();
  } else {
    popUpMessage("loadFail", "지원하지 않는 표 이벤트입니다.");
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
  const source = {
    top: action.position.top,
    left: action.position.left,
    height: action.size.height,
    width: action.size.width,
  };

  Object.assign(chart, source);

  await context.sync();
}

async function applyTableAdded(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.tables.add(action.address);
  await context.sync();
}

export { manageRecording, macroPlay };
