import {
  extractAddresses,
  splitCellAddress,
  getCellsInRange,
} from "./cellCommonUtils";
import ProgressGraph from "./ProgressGraph";

async function parseFormulaSteps() {
  return Excel.run(async (context) => {
    try {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "formulas", "values"]);
      await context.sync();

      const formula = range.formulas[0][0];

      if (!formula) return [];

      const steps = await parseNestedFormula(context, formula);

      await context.sync();

      const validSteps = steps.filter((step) => step !== undefined);
      const sortedSteps = sortStepsByCalculationOrder(validSteps);

      return sortedSteps;
    } catch (e) {
      return [];
    }
  });
}

async function parseNestedFormula(context, formula) {
  const steps = [];
  const stack = [];
  const regex = /([A-Z]+)\(/gi;
  let match;

  while ((match = regex.exec(formula)) !== null) {
    const funcName = match[1];
    const startIndex = regex.lastIndex;
    const args = getArgs(formula, startIndex);

    if (!["DATE", "YEAR", "MONTH", "DAY"].includes(funcName.toUpperCase())) {
      const step = await processFunction(context, funcName, args);

      if (step) {
        if (stack.length > 0) {
          step.dependencies = [stack[stack.length - 1]];
        }
        steps.push(step);
        stack.push(step);
      }
    }
  }

  return steps.reverse();
}

async function processFunction(context, funcName, args) {
  const argList = [];
  const argRegex =
    /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?|\d+(\.\d+)?|"[^"]*"|TRUE|FALSE|[^,()]+)/g;
  let argMatch;

  if (["DATE", "YEAR", "MONTH", "DAY"].includes(funcName.toUpperCase())) {
    return null;
  }

  while ((argMatch = argRegex.exec(args)) !== null) {
    argList.push(argMatch[0]);
  }

  const argValues = [];
  const cellAddresses = new Set();
  let sheet;
  let cell;
  let cellLoadedSuccessfully = true;

  for (const arg of argList) {
    if (arg.match(/((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)/)) {
      const parts = arg.split("!");
      const sheetName =
        parts.length > 1 ? parts[0].replace(/'/g, "") : undefined;
      const normalizedAddress = parts[parts.length - 1].replace(/\$/g, "");

      if (normalizedAddress.includes(":")) {
        const [startCell, endCell] = normalizedAddress.split(":");
        const cellsInRange = getCellsInRange(startCell, endCell);

        cellsInRange.forEach((cells) => cellAddresses.add(cells));
      } else {
        cellAddresses.add(normalizedAddress);
      }

      try {
        sheet = sheetName
          ? context.workbook.worksheets.getItem(sheetName)
          : context.workbook.worksheets.getActiveWorksheet();
      } catch (e) {
        sheet = context.workbook.worksheets.getActiveWorksheet();
      }

      try {
        cell = sheet.getRange(normalizedAddress);

        cell.load(["address", "values", "numberFormat"]);
        await context.sync();
      } catch (e) {
        argValues.push({ arg, value: null });

        cellLoadedSuccessfully = false;
      }

      if (cellLoadedSuccessfully && cell) {
        const cellNumberFormat = cell.numberFormat[0][0];
        let cellValue = cell.values[0][0];

        if (
          cellNumberFormat &&
          cellNumberFormat.includes("yy") &&
          cellValue !== ""
        ) {
          cellValue = new Date((cellValue - 25569) * 86400 * 1000);
        }

        argValues.push({ arg, value: cellValue });
      }
    } else {
      argValues.push({ arg, value: arg });
    }
  }

  const groupedAddresses = groupCellsIntoRanges(Array.from(cellAddresses));
  const formulaOrderInfo = {
    address: groupedAddresses.join(", "),
    functionName: funcName,
    formula: `${funcName}(${args})`,
    dependencies: [],
  };

  applyFunctionSpecificLogic(funcName, args, formulaOrderInfo);

  return formulaOrderInfo;
}

function groupCellsIntoRanges(cells) {
  if (cells.length === 0) return [];

  const ranges = [];
  let startCell = cells[0];
  let endCell = cells[0];

  for (let i = 1; i < cells.length; i += 1) {
    const currentCell = cells[i];

    if (isAdjacent(endCell, currentCell)) {
      endCell = currentCell;
    } else {
      ranges.push(
        startCell === endCell ? startCell : `${startCell}:${endCell}`,
      );

      startCell = currentCell;
      endCell = currentCell;
    }
  }

  ranges.push(startCell === endCell ? startCell : `${startCell}:${endCell}`);

  return ranges;
}

function isAdjacent(cell1, cell2) {
  try {
    const [col1, row1] = splitCellAddress(cell1);
    const [col2, row2] = splitCellAddress(cell2);

    if (col1 === col2 && row2 === row1 + 1) {
      return true;
    }

    if (row1 === row2 && col2.charCodeAt(0) === col1.charCodeAt(0) + 1) {
      return true;
    }

    return false;
  } catch (e) {
    return false;
  }
}

function getArgs(formula, startIndex) {
  let depth = 1;
  let currentArg = "";
  let inString = false;

  for (let i = startIndex; i < formula.length; i += 1) {
    const char = formula[i];

    if (char === '"' && formula[i - 1] !== "\\") {
      inString = !inString;
    }

    if (!inString) {
      if (char === "(") {
        depth += 1;
      } else if (char === ")") {
        depth -= 1;
      }
    }

    if (depth === 0) {
      return currentArg;
    }

    currentArg += char;
  }

  return currentArg;
}

function splitArgs(args) {
  const result = [];
  let depth = 0;
  let currentArg = "";

  for (let i = 0; i < args.length; i += 1) {
    if (args[i] === "," && depth === 0) {
      result.push(currentArg.trim());

      currentArg = "";
    } else {
      currentArg += args[i];

      if (args[i] === "(") {
        depth += 1;
      } else if (args[i] === ")") {
        depth -= 1;
      }
    }
  }

  result.push(currentArg.trim());

  return result;
}

function sortStepsByCalculationOrder(steps) {
  const graph = new ProgressGraph();

  steps.forEach((step) => {
    if (step) {
      graph.addNode(step);
    }
  });

  steps.forEach((step) => {
    if (step && step.dependencies) {
      step.dependencies.forEach((dep) => {
        graph.addDependency(dep, step);
      });
    }
  });

  return graph.topologicalSort();
}

function extractAddressesFromStep(step) {
  const { formula } = step;

  return extractAddresses(formula);
}

function applyFunctionSpecificLogic(funcName, args, formulaOrderInfo) {
  switch (funcName.toUpperCase()) {
    case "IF": {
      const [ifCondition, ifTrueValue, ifFalseValue] = splitArgs(args);
      formulaOrderInfo.condition = ifCondition;
      formulaOrderInfo.trueValue = ifTrueValue;
      formulaOrderInfo.falseValue = ifFalseValue;

      break;
    }
    case "IFS": {
      const ifsArgs = splitArgs(args);
      formulaOrderInfo.conditions = ifsArgs.filter((_, i) => i % 2 === 0);
      formulaOrderInfo.values = ifsArgs.filter((_, i) => i % 2 !== 0);

      break;
    }
    case "SUMIF": {
      const [sumifRange, sumifCriteria, sumifSumRange] = splitArgs(args);
      formulaOrderInfo.criteriaRange = sumifRange;
      formulaOrderInfo.criteria = sumifCriteria;
      formulaOrderInfo.sumRange = sumifSumRange;

      break;
    }
    case "SUMIFS":
    case "COUNTIFS":
    case "AVERAGEIFS": {
      const sumifsArgs = splitArgs(args);
      [formulaOrderInfo.sumRange] = sumifsArgs;
      formulaOrderInfo.criteriaRanges = sumifsArgs
        .slice(1)
        .filter((_, i) => i % 2 === 0);
      formulaOrderInfo.criteria = sumifsArgs
        .slice(1)
        .filter((_, i) => i % 2 !== 0);

      break;
    }
    case "COUNTIF":
    case "AVERAGEIF": {
      const [countifRange, countifCriteria] = splitArgs(args);
      formulaOrderInfo.criteriaRange = countifRange;
      formulaOrderInfo.criteria = countifCriteria;

      break;
    }
    case "IFERROR":
    case "IFNA": {
      const [ifErrorValue, ifErrorAlternative] = splitArgs(args);
      formulaOrderInfo.condition = ifErrorValue;
      formulaOrderInfo.falseValue = ifErrorAlternative;

      break;
    }
    case "SWITCH": {
      const [switchExpression, ...switchArgs] = splitArgs(args);
      formulaOrderInfo.condition = switchExpression;
      formulaOrderInfo.values = switchArgs;

      break;
    }
    case "CHOOSE": {
      const [chooseIndex, ...chooseValues] = splitArgs(args);
      formulaOrderInfo.condition = chooseIndex;
      formulaOrderInfo.values = chooseValues;

      break;
    }
    default: {
      return {};
    }
  }
  return formulaOrderInfo;
}

export {
  processFunction,
  parseFormulaSteps,
  parseNestedFormula,
  groupCellsIntoRanges,
  isAdjacent,
  getArgs,
  splitArgs,
  extractAddressesFromStep,
};
