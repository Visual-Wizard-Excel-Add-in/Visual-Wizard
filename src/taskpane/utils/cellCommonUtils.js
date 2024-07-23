import useStore from "./store";

function updateState(setStateFunc, newValue) {
  useStore.getState()[setStateFunc](newValue);
}

function splitCellAddress(address) {
  const match = address.match(/\$?([A-Z]+)\$?([0-9]+)/);

  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  return [match[1], parseInt(match[2], 10)];
}

function extractAddresses(arg) {
  const argAddresses = [];
  const argRegex = /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)/g;
  let match;

  while ((match = argRegex.exec(arg)) !== null) {
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

function extractArgsAddress(cellArgument) {
  const cleanedArgument = cellArgument.replace(/\$/g, "");
  const match = cleanedArgument.match(/([A-Z]+\d+)/);

  return match ? match[1] : null;
}

async function getCellValue() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "formulas", "values", "numberFormat"]);
      await context.sync();

      const selectedCellAddress = range.address;
      const numberFormat = range.numberFormat[0][0];
      let selectedCellValue = range.values[0][0];
      let formula = range.formulas[0][0];

      if (
        !range ||
        !range.values ||
        !range.values[0] ||
        range.values[0][0] === undefined
      ) {
        return;
      }

      if (typeof formula !== "string") {
        formula = "";
      }

      const formulaFunctions = extractFunctionsFromFormula(formula);
      const formulaArgs = await extractArgsFromFormula(formula);

      if (
        numberFormat &&
        numberFormat.includes("yy") &&
        selectedCellValue !== ""
      ) {
        selectedCellValue = new Date(
          (selectedCellValue - 25569) * 86400 * 1000,
        ).toLocaleDateString();
      }

      updateState("setCellAddress", selectedCellAddress);
      updateState("setCellValue", selectedCellValue);
      updateState("setCellFormula", range.formulas[0][0]);
      updateState("setCellFunctions", formulaFunctions);
      updateState("setCellArguments", formulaArgs);

      await context.sync();
    });
  } catch (e) {
    throw new Error(e.message);
  }
}

async function getTargetCellValue(targetCell) {
  return Excel.run(async (context) => {
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

    if (!cell.values || !cell.values[0] || cell.values[0][0] === undefined) {
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
}

async function extractArgsFromFormula(formula) {
  const argSet = new Set();
  const argCellRegex = /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)/g;
  const promises = [];
  const matches = formula.match(argCellRegex);

  if (matches) {
    for (const matchedArg of matches) {
      if (matchedArg.includes(":")) {
        const [startCell, endCell] = matchedArg.split(":");
        const cellsInRange = getCellsInRange(startCell, endCell);

        for (const cell of cellsInRange) {
          if (!argSet.has(cell)) {
            argSet.add(cell);

            const value = await getTargetCellValue(cell);

            promises.push(`${cell}(${value})`);
          }
        }
      } else if (!argSet.has(matchedArg)) {
        argSet.add(matchedArg);

        const value = await getTargetCellValue(matchedArg);

        promises.push(`${matchedArg}(${value})`);
      }
    }
  }

  return Promise.all(promises);
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

async function registerSelectionChange(sheetId, func) {
  await Excel.run(async (context) => {
    const { workbook } = context;
    const sheet = workbook.worksheets.getItem(sheetId);

    sheet.onSelectionChanged.add(func);
    await context.sync();
  });
}

async function activeSheetId(sheetId) {
  await Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();

    activeSheet.load("name");
    await context.sync();

    const activatedSheetId = activeSheet.name;

    if (activatedSheetId !== sheetId) {
      updateState("setSheetName", activatedSheetId);
      await registerSelectionChange(context, activatedSheetId, getCellValue);
    }
  });
}

async function addPreset(presetCategory, presetName) {
  await Excel.run(async () => {
    let savePreset = Office.context.document.settings.get(presetCategory);

    if (!savePreset) {
      savePreset = {};
    } else {
      savePreset = JSON.parse(savePreset);
    }

    savePreset[presetName] = {};

    Office.context.document.settings.set(
      presetCategory,
      JSON.stringify(savePreset),
    );
    await Office.context.document.settings.saveAsync();
  });
}

async function deletePreset(presetCategory, presetName) {
  await Excel.run(async () => {
    let currentPresets = Office.context.document.settings.get(presetCategory);

    if (currentPresets) {
      currentPresets = JSON.parse(currentPresets);

      delete currentPresets[presetName];

      Office.context.document.settings.set(
        presetCategory,
        JSON.stringify(currentPresets),
      );
      await Office.context.document.settings.saveAsync();
    }
  });
}

export {
  registerSelectionChange,
  getCellValue,
  updateState,
  splitCellAddress,
  extractAddresses,
  extractArgsAddress,
  getCellsInRange,
  nextColumn,
  activeSheetId,
  addPreset,
  deletePreset,
};
