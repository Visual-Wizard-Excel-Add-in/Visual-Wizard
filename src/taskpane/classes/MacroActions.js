import { selectRangeValues, popUpMessage } from "../utils/commonFuncs";
import { extractCellStyle } from "../utils/cellStyleFuncs";

class MacroAction {
  constructor(event) {
    this.type = event.type;
    this.init(event);
  }

  async init(event) {
    switch (this.type) {
      case "WorksheetChanged":
        this.address = event.address;
        this.details = {
          value: event.details
            ? event.details.valueAfter
            : await selectRangeValues(),
        };
        break;

      case "WorksheetFormatChanged":
        this.address = event.address;
        this.cellStyle = await recordCellStyle(this.address);
        break;

      case "TableChanged":
        this.address = event.address;
        this.details = event.details;
        this.tableId = event.tableId;
        this.changeType = event.changeType;
        break;

      case "ChartAdded":
        this.chartId = event.chartId;
        await onChartAdded(this);
        break;

      case "TableAdded":
        this.tableId = event.tableId;
        [this.address, this.showHeaders] = await onTableAdded(this);
        break;

      default:
        popUpMessage("loadFail", "지원하지 않는 형식입니다.");
        break;
    }
  }
}

export default MacroAction;

async function recordCellStyle(address) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(address);

    return await extractCellStyle(context, cell);
  });
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
