import { updateState } from "./commonFuncs";

async function getLastCellAddress() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();

    usedRange.load(["values"]);
    await context.sync();

    let lastRowIndex = -1;
    let lastColIndex = -1;

    usedRange.values.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        if (value !== "") {
          if (rowIndex > lastRowIndex) lastRowIndex = rowIndex;

          if (colIndex > lastColIndex) lastColIndex = colIndex;
        }
      });
    });

    if (lastRowIndex === -1 || lastColIndex === -1) {
      return null;
    }

    const lastCell = usedRange.getCell(lastRowIndex, lastColIndex);

    lastCell.load("address");
    await context.sync();

    const lastCellAddress = lastCell.address.split("!")[1];

    return lastCellAddress;
  });
}

async function evaluateTestFormula(newFormula) {
  try {
    let testResult = "";

    await Excel.run(async (context) => {
      const { workbook } = context;
      const originSheet = workbook.worksheets.getActiveWorksheet();

      try {
        workbook.worksheets.getItem("TestSheet").delete();
        await context.sync();
      } catch (error) {
        if (error.code !== Excel.ErrorCodes.itemNotFound) {
          throw error;
        }
      }

      originSheet.load("name");
      await context.sync();

      const originSheetName = originSheet.name;
      const testSheet = workbook.worksheets.add("TestSheet");
      const sheetRefFormula = newFormula
        .split(",")
        .map((segment) => {
          return segment.replace(
            /((?:[^!]+!)?\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/g,
            (match) => {
              if (match.includes("!")) {
                return match;
              }
              return `${originSheetName}!${match}`;
            },
          );
        })
        .join(",");

      const formulaRange = testSheet.getRange("A1");

      formulaRange.formulas = [[sheetRefFormula]];
      formulaRange.load("values");
      await context.sync();

      [[testResult]] = formulaRange.values;

      testSheet.delete();
      await context.sync();
    });

    return testResult;
  } catch (e) {
    const warningMessage = {
      type: "warning",
      title: "에러 발생: ",
      body: `테스트를 진행 중 에러가 발생했습니다.${e.message}`,
    };

    updateState("setMessageList", warningMessage);

    return null;
  }
}

export { getLastCellAddress, evaluateTestFormula };
