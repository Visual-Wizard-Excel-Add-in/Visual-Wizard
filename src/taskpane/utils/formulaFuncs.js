import ProgressGraph from "../classes/ProgressGraph";

async function parseFormulaSteps(formula) {
  try {
    if (!formula) return [];

    const steps = await parseNestedFormula(formula);
    const validSteps = steps.filter((step) => step !== undefined);
    const sortedSteps = sortCalculationOrder(validSteps);

    return sortedSteps;
  } catch (e) {
    return [];
  }
}

async function parseNestedFormula(formula) {
  const steps = [];
  const stack = [];
  const regex = /([A-Z]+)\(/g;
  let match;

  while ((match = regex.exec(formula)) !== null) {
    const funcName = match[1];
    const argsStartIndex = regex.lastIndex;
    const args = getArgs(formula, argsStartIndex);

    if (!["DATE", "YEAR", "MONTH", "DAY"].includes(funcName.toUpperCase())) {
      const step = await getFormulaDetail(funcName, args);

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

async function getFormulaDetail(funcName, args) {
  const argList = [];
  const cellAddresses = new Set();
  const referenceRegex =
    /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?|\d+(\.\d+)?|"[^"]*"|TRUE|FALSE|[^,()]+)/g;
  const onlyCellRegex =
    /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)/;
  let argMatch = null;

  if (["DATE", "YEAR", "MONTH", "DAY"].includes(funcName.toUpperCase())) {
    return null;
  }

  const argsArray = args.split(",");

  for (const arg of argsArray) {
    while ((argMatch = referenceRegex.exec(arg.trim())) !== null) {
      argList.push(argMatch[0]);
    }
  }

  for (const arg of argList) {
    if (arg.match(onlyCellRegex)) {
      const parts = arg.split("!");
      const normalizedAddress = parts[parts.length - 1].replace(/\$/g, "");

      cellAddresses.add(normalizedAddress);
    }
  }

  const formulaStepInfo = {
    address: Array.from(cellAddresses).join(", "),
    functionName: funcName,
    formula: `${funcName}(${args})`,
    dependencies: [],
  };

  getConditionFuncInfo(funcName, args, formulaStepInfo);

  return formulaStepInfo;
}

function groupCellsIntoRanges(cells) {
  if (cells.length === 0) return [];

  const sortedCells = cells.sort((a, b) => {
    const colA = a.match(/[A-Z]+/)[0];
    const colB = b.match(/[A-Z]+/)[0];
    const rowA = parseInt(a.match(/\d+/)[0], 10);
    const rowB = parseInt(b.match(/\d+/)[0], 10);

    if (colA === colB) {
      return rowA - rowB;
    }

    return colA.localeCompare(colB);
  });

  const ranges = [];
  let startCell = sortedCells[0];
  let endCell = sortedCells[0];

  for (let i = 1; i < sortedCells.length; i += 1) {
    const currentCell = sortedCells[i];
    const [currentColumn, currentRow] = [
      currentCell.match(/[A-Z]+/)[0],
      parseInt(currentCell.match(/\d+/)[0], 10),
    ];
    const [endColumn, endRow] = [
      endCell.match(/[A-Z]+/)[0],
      parseInt(endCell.match(/\d+/)[0], 10),
    ];

    if (
      (currentColumn === endColumn && currentRow === endRow + 1) ||
      (currentRow === endRow &&
        currentColumn.charCodeAt(0) === endColumn.charCodeAt(0) + 1)
    ) {
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

  const individualCells = ranges.filter((range) => !range.includes(":"));
  const rangeCells = ranges.filter((range) => range.includes(":"));

  const mergedRanges = [];
  let currentRange = rangeCells.length > 0 ? rangeCells[0] : null;

  for (let i = 1; i < rangeCells.length; i += 1) {
    const [startRange, endRange] = currentRange.split(":");
    const [nextStartRange, nextEndRange] = rangeCells[i].split(":");

    const currentStartCol = startRange.match(/[A-Z]+/)[0];
    const currentEndCol = endRange.match(/[A-Z]+/)[0];
    const nextStartCol = nextStartRange.match(/[A-Z]+/)[0];
    const nextEndCol = nextEndRange.match(/[A-Z]+/)[0];

    const currentStartRow = parseInt(startRange.match(/\d+/)[0], 10);
    const currentEndRow = parseInt(endRange.match(/\d+/)[0], 10);
    const nextStartRow = parseInt(nextStartRange.match(/\d+/)[0], 10);
    const nextEndRow = parseInt(nextEndRange.match(/\d+/)[0], 10);

    if (
      currentEndCol.charCodeAt(0) + 1 === nextStartCol.charCodeAt(0) &&
      nextStartRow >= currentStartRow &&
      nextEndRow <= currentEndRow
    ) {
      currentRange = `${currentStartCol}${currentStartRow}:${nextEndCol}${nextEndRow}`;
    } else {
      mergedRanges.push(currentRange);

      currentRange = rangeCells[i];
    }
  }

  if (currentRange) {
    mergedRanges.push(currentRange);
  }

  return [...individualCells, ...mergedRanges];
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

function sortCalculationOrder(steps) {
  const orderGraph = new ProgressGraph();

  steps.forEach((step) => {
    if (step) {
      orderGraph.addNode(step);
    }
  });

  steps.forEach((step) => {
    if (step && step.dependencies) {
      step.dependencies.forEach((dep) => {
        orderGraph.addDependency(dep, step);
      });
    }
  });

  return orderGraph.topologicalSort();
}

function getConditionFuncInfo(funcName, args, formulaOrderInfo) {
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
  getFormulaDetail as processFunction,
  parseFormulaSteps,
  parseNestedFormula,
  groupCellsIntoRanges,
  getArgs,
  splitArgs,
};
