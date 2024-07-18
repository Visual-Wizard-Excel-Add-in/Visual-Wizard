async function storeCellStyle(
  context,
  cell,
  allCellStyles,
  isCellHighlighting,
) {
  const edges = ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"];

  cell.load(["address", "format/fill"]);

  await context.sync();

  if (allCellStyles[cell.address] && !isCellHighlighting) {
    return;
  }

  const borders = {};

  for (const edge of edges) {
    borders[edge] = cell.format.borders.getItem(edge);

    borders[edge].load(["color", "style", "weight"]);
  }
  await context.sync();

  const cellStyle = {
    color: cell.format.fill.color,
    borders: {},
  };

  for (const edge of edges) {
    const border = borders[edge];

    cellStyle.borders[edge] = {
      color: border.color,
      style: border.style,
      weight: border.weight,
    };
  }

  if (!allCellStyles[cell.address] && isCellHighlighting) {
    allCellStyles[cell.address] = cellStyle;
  }

  Office.context.document.settings.set("allCellStyles", allCellStyles);

  await Office.context.document.settings.saveAsync();
}

async function applyCellStyle(
  context,
  cell,
  allCellStyles,
  isCellHighlighting,
) {
  const edges = ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"];

  cell.load(["address"]);

  await context.sync();

  const cellStyle = allCellStyles[cell.address];

  if (cellStyle && !isCellHighlighting) {
    cell.format.fill.color = cellStyle.color;

    for (const edge of edges) {
      if (Object.prototype.hasOwnProperty.call(cellStyle.borders, edge)) {
        const border = cell.format.borders.getItem(edge);
        const borderStyle = cellStyle.borders[edge];

        if (borderStyle.style !== "None") {
          border.color = borderStyle.color;
          border.style = borderStyle.style;
          border.weight = borderStyle.weight;
        } else {
          border.style = Excel.BorderLineStyle.continuous;
          border.color = "#d6d6d6";
          border.weight = "thin";
        }
      }
    }

    await context.sync();

    delete allCellStyles[cell.address];

    Office.context.document.settings.set("allCellStyles", allCellStyles);

    await Office.context.document.settings.saveAsync();
  }
}

async function highlightingCell(isCellHighlighting, argCells, resultCell) {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const resultCellRange = worksheet.getRange(resultCell);
    const allCellStyles =
      Office.context.document.settings.get("allCellStyles") || {};

    context.trackedObjects.add(resultCellRange);

    await storeCellStyle(
      context,
      resultCellRange,
      allCellStyles,
      isCellHighlighting,
    );

    const argsCellAddresses = argCells.map(extractArgsAddress).filter(Boolean);

    for (const argcell of argsCellAddresses) {
      const argcellsRange = worksheet.getRange(argcell);

      context.trackedObjects.add(argcellsRange);

      await storeCellStyle(
        context,
        argcellsRange,
        allCellStyles,
        isCellHighlighting,
      );

      context.trackedObjects.remove(argcellsRange);
    }

    await context.sync();

    if (isCellHighlighting) {
      resultCellRange.format.fill.color = "#3d33ff";

      await ChangeCellborder(resultCellRange, "red", false);
    } else {
      await applyCellStyle(
        context,
        resultCellRange,
        allCellStyles,
        isCellHighlighting,
      );
    }

    context.trackedObjects.remove(resultCellRange);

    for (const argcell of argsCellAddresses) {
      const argcellsRange = worksheet.getRange(argcell);

      context.trackedObjects.add(argcellsRange);

      if (isCellHighlighting) {
        argcellsRange.format.fill.color = "#28f925";

        await ChangeCellborder(argcellsRange, "red", false);
      } else {
        await applyCellStyle(
          context,
          argcellsRange,
          allCellStyles,
          isCellHighlighting,
        );
      }
      context.trackedObjects.remove(argcellsRange);
    }
    await context.sync();
  });
}

function extractArgsAddress(cellArgument) {
  const cleanedArgument = cellArgument.replace(/\$/g, "");
  const match = cleanedArgument.match(/([A-Z]+\d+)/);

  return match ? match[1] : null;
}

async function getCellValue(
  setCellArguments,
  setCellAddress,
  setCellValue,
  setCellFormulas,
) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();

    range.load(["address", "formulas", "values", "numberFormat"]);

    await context.sync();

    let selectedCellValue = range.values[0][0];
    const formulaFunctions = extractFunctionsFromFormula(range.formulas[0][0]);
    const numberFormat = range.numberFormat[0][0];
    const formulaArgs = await extractArgsFromFormula(range.formulas[0][0]);
    const selectedCellAddress = range.address.match(/\$?[A-Z]+\$?[0-9]+/g)[0];

    if (
      numberFormat &&
      numberFormat.includes("yy") &&
      selectedCellValue !== ""
    ) {
      selectedCellValue = new Date(
        (selectedCellValue - 25569) * 86400 * 1000,
      ).toLocaleDateString();
    }

    setCellAddress(selectedCellAddress);
    setCellValue(selectedCellValue);
    setCellFormulas(formulaFunctions);
    setCellArguments(formulaArgs);
  });
}

async function getTargetCellValue(targetCell) {
  return Excel.run(async (context) => {
    const cell = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(targetCell);

    cell.load(["values", "numberFormat"]);

    await context.sync();

    let targetCellValue = cell.values[0][0];
    const numberFormat = cell.numberFormat[0][0];

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
  const argRegex = /\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?/g;
  const promises = [];
  const matches = formula.match(argRegex);

  if (matches) {
    matches.forEach((matchedArg) => {
      if (matchedArg.includes(":")) {
        const [startCell, endCell] = matchedArg.split(":");
        const cellsInRange = getCellsInRange(startCell, endCell);

        cellsInRange.forEach((cell) => {
          if (!argSet.has(cell)) {
            argSet.add(cell);

            const promise = getTargetCellValue(cell).then((value) => {
              return `${cell}(${value})`;
            });

            promises.push(promise);
          }
        });
      } else if (!argSet.has(matchedArg)) {
        argSet.add(matchedArg);

        const promise = getTargetCellValue(matchedArg).then((value) => {
          return `${matchedArg}(${value})`;
        });

        promises.push(promise);
      }
    });
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

async function ChangeCellborder(targetCell, color, isClear) {
  const edges = ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"];

  for (const edge of edges) {
    const border = targetCell.format.borders.getItem(edge);

    if (isClear) {
      border.style = Excel.BorderLineStyle.none;
    } else {
      border.color = color;
      border.style = Excel.BorderLineStyle.continuous;
      border.weight = Excel.BorderWeight.thick;
    }
  }
  await targetCell.context.sync();
}

async function registerSelectionChange(func) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.onSelectionChanged.add(func);
  });
}

export { getCellValue, registerSelectionChange, highlightingCell };
