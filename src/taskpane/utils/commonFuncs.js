import CellInfo from "../classes/CellInfo";
import Message from "../classes/Message";
import usePublicStore from "../store/publicStore";

function updateState(setStateFunc, newValue) {
  usePublicStore.getState()[setStateFunc](newValue);
}

async function updateCellInfo() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "formulas", "values", "numberFormat"]);
      await context.sync();

      const selectCell = new CellInfo(
        range.address,
        range.values[0][0],
        range.formulas[0][0],
        range.numberFormat[0][0],
      );

      const formulaFunctions = extractFunctionsFromFormula(selectCell.formula);
      const formulaArgs = await extractArgsFromFormula(selectCell.formula);

      updateCellState();

      function updateCellState() {
        const stateMapping = {
          cellAddress: { value: selectCell.address, setter: "setCellAddress" },
          cellValue: { value: selectCell.values, setter: "setCellValue" },
          cellFormula: { value: selectCell.formula, setter: "setCellFormula" },
          cellFunctions: {
            value: formulaFunctions,
            setter: "setCellFunctions",
          },
          cellArgument: { value: formulaArgs, setter: "setCellArguments" },
        };

        Object.keys(stateMapping).forEach((state) => {
          const { value, setter } = stateMapping[state];

          if (isChanged(value, usePublicStore.getState()[state])) {
            updateState(setter, value);
          }
        });

        function isChanged(cellValue, state) {
          return cellValue !== state;
        }
      }
    });
  } catch (error) {
    throw new Error(error.message);
  }
}

async function getTargetCellValue(targetCell) {
  const targetValue = await Excel.run(async (context) => {
    const parts = targetCell.split("!");
    const sheetName = parts.length > 1 ? parts[0] : undefined;
    const normalizedAddress =
      parts.length > 1
        ? parts[1].replace(/\$/g, "")
        : parts[0].replace(/\$/g, "");

    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(normalizedAddress);

    cell.load(["values", "numberFormat"]);
    await context.sync();

    if (cell.values[0][0] === "") {
      return null;
    }

    const numberFormat = cell.numberFormat[0][0];
    let targetCellValue = cell.values[0][0];

    if (numberFormat && numberFormat.includes("yy") && targetCellValue !== "") {
      targetCellValue = new Date(
        (targetCellValue - 25569) * 86400 * 1000,
      ).toLocaleDateString();
    }

    return targetCellValue;
  });

  return targetValue;
}

async function getSelectRangeValue() {
  let rangeValue = null;

  await Excel.run(async (context) => {
    const selectRange = context.workbook.getSelectedRange();

    selectRange.load("values");
    await context.sync();

    rangeValue = selectRange.values;
  });

  return rangeValue;
}

let handleSelectionChange = null;

async function registerSelectionChange(sheetId, func) {
  if (handleSelectionChange !== null) {
    await Excel.run(handleSelectionChange.context, async (context) => {
      handleSelectionChange?.remove();
      await context.sync();

      handleSelectionChange = null;
    });
  }

  await Excel.run(async (context) => {
    const { workbook } = context;
    const sheet = workbook.worksheets.getItem(sheetId);

    handleSelectionChange = sheet.onSelectionChanged.add(func);
    await context.sync();
  });
}

function splitCellAddress(address) {
  const match = address.match(/\$?([A-Z]+)\$?([0-9]+)/);

  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  return [match[1], parseInt(match[2], 10)];
}

function extractReferenceCells(formula) {
  const argAddresses = [];
  const argRegex = /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)/g;
  let match;

  while ((match = argRegex.exec(formula)) !== null) {
    const parts = match[0].split("!");
    const normalizedAddress = parts[parts.length - 1].replace(/\$/g, "");

    if (normalizedAddress.includes(":")) {
      const [startCell, endCell] = normalizedAddress.split(":");
      const cellsInRange = getCellsInRange(startCell, endCell);

      argAddresses.push(...cellsInRange);
    } else {
      argAddresses.push(normalizedAddress);
    }
  }

  return argAddresses;
}

async function extractArgsFromFormula(formula) {
  const argSet = new Set();
  const argCellRegex =
    /([A-Z]+[0-9]+|\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)/g;
  const results = [];
  const matches = formula.match(argCellRegex);

  if (matches) {
    for (const matchedArg of matches) {
      if (matchedArg.includes(":")) {
        const [startCell, endCell] = matchedArg.split(":");
        const cellsInRange = getCellsInRange(startCell, endCell);

        for (const cell of cellsInRange) {
          if (!argSet.has(cell)) {
            const value = await getTargetCellValue(cell);

            argSet.add(cell);
            results.push(`${cell}(${value})`);
          }
        }
      } else if (!argSet.has(matchedArg)) {
        const value = await getTargetCellValue(matchedArg);

        argSet.add(matchedArg);
        results.push(`${matchedArg}(${value})`);
      }
    }
  }

  return results;
}

function extractFunctionsFromFormula(formula) {
  const functionNames = [];
  const regex = /([A-Z]+)\(/g;
  let match;

  while ((match = regex.exec(formula)) !== null) {
    if (!functionNames.includes(match[1])) {
      functionNames.push(match[1]);
    }
  }

  return functionNames;
}

function getCellsInRange(startCell, endCell) {
  const cells = [];
  const startColumn = startCell.match(/[A-Z]+/)[0];
  const startRow = parseInt(startCell.match(/[0-9]+/)[0], 10);
  const endColumn = endCell.match(/[A-Z]+/)[0];
  const endRow = parseInt(endCell.match(/[0-9]+/)[0], 10);
  let currentColumn = startColumn;

  while (currentColumn <= endColumn) {
    for (let row = startRow; row <= endRow; row += 1) {
      cells.push(`${currentColumn}${row}`);
    }

    if (currentColumn === endColumn) {
      break;
    }

    currentColumn = nextColumn(currentColumn);
  }

  return cells;
}

function nextColumn(col) {
  if (col === "Z") {
    return "AA";
  }

  if (col.length === 1) {
    return String.fromCharCode(col.charCodeAt(0) + 1);
  }

  let lastChar = col.slice(-1);
  let restChars = col.slice(0, -1);

  if (lastChar === "Z") {
    restChars = nextColumn(restChars);
    lastChar = "A";
  } else {
    lastChar = String.fromCharCode(lastChar.charCodeAt(0) + 1);
  }

  return restChars + lastChar;
}

async function addPreset(presetCategory, presetName) {
  let savePreset = await OfficeRuntime.storage.getItem(presetCategory);

  if (!savePreset) {
    savePreset = {};
  } else {
    savePreset = JSON.parse(savePreset);
  }

  savePreset[presetName] = {};

  await OfficeRuntime.storage.setItem(
    presetCategory,
    JSON.stringify(savePreset),
  );
}

async function deletePreset(presetCategory, presetName) {
  let currentPresets = await OfficeRuntime.storage.getItem(presetCategory);

  if (currentPresets) {
    currentPresets = JSON.parse(currentPresets);

    delete currentPresets[presetName];

    await OfficeRuntime.storage.setItem(
      presetCategory,
      JSON.stringify(currentPresets),
    );
  }
}

function popUpMessage(purpose = null, option = "") {
  updateState("setMessageList", new Message(purpose, option).body);
}

export {
  registerSelectionChange,
  updateCellInfo,
  getSelectRangeValue,
  updateState,
  splitCellAddress,
  extractReferenceCells,
  getCellsInRange,
  nextColumn,
  addPreset,
  deletePreset,
  getTargetCellValue,
  extractArgsFromFormula,
  extractFunctionsFromFormula,
  popUpMessage,
};
