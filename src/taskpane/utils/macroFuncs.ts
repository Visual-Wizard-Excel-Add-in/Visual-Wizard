import { storeCellStyle, applyCellStyle } from "./cellStyleFunc";
import { updateState } from "./cellCommonUtils";

let worksheetChangedHandler: OfficeExtension.EventHandlerResult<Excel.WorksheetChangedEventArgs> | null;
let tableChangedHandler: OfficeExtension.EventHandlerResult<Excel.TableChangedEventArgs> | null;
let chartAddedHandler: OfficeExtension.EventHandlerResult<Excel.ChartAddedEventArgs> | null;
let tableAddedHandler: OfficeExtension.EventHandlerResult<Excel.TableAddedEventArgs> | null;
let formatChangedHandler: OfficeExtension.EventHandlerResult<Excel.WorksheetFormatChangedEventArgs> | null;
let actions: MacroActionType[] = [];

async function manageRecording(isRecording: boolean, presetName: string) {
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

      let allMacroPresets = await JSON.parse(
        OfficeRuntime.storage.getItem("allMacroPresets"),
      );

      allMacroPresets[presetName] = { actions: [] };

      await OfficeRuntime.storage.setItem(
        "allMacroPresets",
        JSON.stringify(allMacroPresets),
      );
      await context.sync();
    } else {
      if (worksheetChangedHandler) {
        await Excel.run(worksheetChangedHandler.context, async (ctx) => {
          if (worksheetChangedHandler !== null) {
            worksheetChangedHandler.remove();
            await ctx.sync();
          }
        });
      }

      if (tableChangedHandler) {
        await Excel.run(tableChangedHandler.context, async (ctx) => {
          if (tableChangedHandler !== null) {
            tableChangedHandler.remove();
            await ctx.sync();
          }
        });
      }

      if (chartAddedHandler) {
        await Excel.run(chartAddedHandler.context, async (ctx) => {
          if (chartAddedHandler !== null) {
            chartAddedHandler.remove();
            await ctx.sync();
          }
        });
      }

      if (tableAddedHandler) {
        await Excel.run(tableAddedHandler.context, async (ctx) => {
          if (tableAddedHandler !== null) {
            tableAddedHandler.remove();
            await ctx.sync();
          }
        });
      }

      if (formatChangedHandler) {
        await Excel.run(formatChangedHandler.context, async (ctx) => {
          if (formatChangedHandler !== null) {
            formatChangedHandler.remove();
            await ctx.sync();
          }
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

async function onWorksheetChanged(
  event:
    | Excel.WorksheetChangedEventArgs
    | Excel.TableAddedEventArgs
    | Excel.TableChangedEventArgs
    | Excel.WorksheetFormatChangedEventArgs
    | Excel.ChartAddedEventArgs,
  presetName: string,
) {
  const action: MacroActionType = { type: event.type };
  let cellStyleData: CellStyleType | null = null;
  let allMacroPresets = await JSON.parse(
    OfficeRuntime.storage.getItem("allMacroPresets"),
  );
  let warningMessage: MessageType = { type: "", title: "", body: "" };

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

        if (cellStyleData !== null) {
          action.cellStyle = cellStyleData;
        }

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

        const tableInfo = await onTableAdded(action);

        if (tableInfo) {
          [action.address, action.showHeaders] = tableInfo;
        }
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
    if (e instanceof Error) {
      warningMessage = {
        type: "error",
        title: "에러 발생: ",
        body: `기록 중 예상치 못한 에러가 발생했습니다. ${e.message}`,
      };
    }

    updateState("setMessageList", warningMessage);
  }

  if (action.chartType === "Invalid") {
    warningMessage = {
      type: "warning",
      title: "지원하지 않는 차트: ",
      body: "매크로 설정에서 직접 차트 타입을 입력해주세요.",
    };

    updateState("setMessageList", warningMessage);
  }

  actions.push(action);
}

async function saveMacro(presetName: string) {
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
    if (error instanceof Error) {
      const warningMessage = {
        type: "warning",
        title: "저장 실패",
        body: `매크로 기록에 실패했습니다. ${error.message}`,
      };

      updateState("setMessageList", warningMessage);
    }
  }
}

async function onChartAdded(action: MacroActionType) {
  let chartId: string;

  if (action.chartId) {
    chartId = action.chartId;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.getItem(chartId);
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
          let valuesDataSource: OfficeExtension.ClientResult<string> | string;

          try {
            valuesDataSource = series.getDimensionDataSourceString("Values");
          } catch (e) {
            valuesDataSource = { value: "Invalid" };
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

      action.chartType = chart.chartType || "Invalid";
      action.position = { top: chart.top, left: chart.left };
      action.size = { height: chart.height, width: chart.width };
      action.dataRange = dataRange.map((range) => range.address);

      return action;
    });
  } catch (e) {
    if (e instanceof Error) {
      const warningMessage = {
        type: "warning",
        title: "저장 실패",
        body: `매크로 기록에 실패했습니다. ${e.message}`,
      };

      updateState("setMessageList", warningMessage);
    }
  }
}

async function onTableAdded(action: MacroActionType) {
  try {
    let tableId: string;
    let tableAttributes: [string, boolean];

    if (action.tableId) {
      tableId = action.tableId;
    }

    tableAttributes = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem(tableId);
      const tableRange = table.getRange();

      tableRange.load("address");
      table.load("showHeaders");
      await context.sync();

      const { showHeaders } = table;

      return (tableAttributes = [
        tableRange.address.split("!")[1],
        showHeaders,
      ]);
    });

    if (tableAttributes) {
      return tableAttributes;
    }
  } catch (e) {
    if (e instanceof Error) {
      const warningMessage = {
        type: "warning",
        title: "저장 실패",
        body: `매크로 기록에 실패했습니다. ${e.message}`,
      };

      updateState("setMessageList", warningMessage);

      return null;
    }
  }
}

async function macroPlay(presetName: string) {
  try {
    await Excel.run(async (context) => {
      const allMacroPresets: string =
        await OfficeRuntime.storage.getItem("allMacroPresets");
      let warningMessage: MessageType = { type: "", title: "", body: "" };

      if (!allMacroPresets) {
        throw new Error("No macros found.");
      }

      const parsedPresets: MacroPresetsType = JSON.parse(allMacroPresets);
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
            if (action.address) {
              await applyCellStyle(
                action.address,
                "allMacroPresets",
                false,
                action.cellStyle,
              );
            }
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

async function applyWorksheetChange(
  context: Excel.RequestContext,
  action: MacroActionType,
) {
  if (action.details && "value" in action.details && action.details.value) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(action.address);

    range.values = [[action.details.value]];
    await context.sync();
  }
}

async function applyTableChange(
  context: Excel.RequestContext,
  action: MacroActionType,
) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let warningMessage = { type: "", title: "", body: "" };

  switch (action.changeType) {
    case "RangeEdited":
      if (
        action.details &&
        !("value" in action.details) &&
        action.details.valueAfter
      ) {
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

async function applyChartAdded(
  context: Excel.RequestContext,
  action: MacroActionType,
) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let chartType: ChartType = "Invalid";
  let mergedRange: string | undefined = undefined;

  if (action.chartType) {
    chartType = action.chartType;
  }

  if (action.dataRange && action.dataRange[0].includes(":")) {
    mergedRange = `${action.dataRange[0].split(":")[0]}:${action.dataRange[action.dataRange.length - 1].split(":")[1]}`;
  } else if (action.dataRange) {
    mergedRange = `${action.dataRange[0]}:${action.dataRange[action.dataRange.length - 1]}`;
  }

  const chart = sheet.charts.add(chartType, sheet.getRange(mergedRange));

  if (action.position) {
    chart.top = action.position.top;
    chart.left = action.position.left;
  }

  if (action.size) {
    chart.height = action.size.height;
    chart.width = action.size.width;
  }

  await context.sync();
}

async function applyTableAdded(
  context: Excel.RequestContext,
  action: MacroActionType,
) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  if (action.address) {
    sheet.tables.add(
      action.address,
      action.showHeaders ? action.showHeaders : true,
    );
  }

  await context.sync();
}

export { manageRecording, macroPlay, saveMacro };
