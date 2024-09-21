import { splitCellAddress, getCellsInRange } from "./cellCommonUtils";
import ProgressGraph from "./ProgressGraph";

interface ArgValueType {
  arg: string;
  value: string | number | Date | null;
}

async function parseFormulaSteps() {
  return Excel.run(async (context) => {
    try {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "formulas", "values"]);
      await context.sync();

      const formula: string = range.formulas[0][0];

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

async function parseNestedFormula(
  context: Excel.RequestContext,
  formula: string,
): Promise<GraphNodeType[]> {
  const steps: GraphNodeType[] = [];
  const stack: GraphNodeType[] | null = [];
  const regex = /([A-Z]+)\(/gi;
  let match: RegExpExecArray | null;

  while ((match = regex.exec(formula)) !== null) {
    const funcName = match[1];
    const startIndex = regex.lastIndex;
    const args = getArgs(formula, startIndex);

    if (!["DATE", "YEAR", "MONTH", "DAY"].includes(funcName.toUpperCase())) {
      const step: GraphNodeType | null = await processFunction(
        context,
        funcName,
        args,
      );

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

async function processFunction(
  context: Excel.RequestContext,
  funcName: string,
  args: string,
) {
  const argList: string[] = [];
  const argRegex =
    /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?|\d+(\.\d+)?|"[^"]*"|TRUE|FALSE|[^,()]+)/g;
  let argMatch;

  if (["DATE", "YEAR", "MONTH", "DAY"].includes(funcName.toUpperCase())) {
    return null;
  }

  const argsArray = args.split(",");

  for (const arg of argsArray) {
    while ((argMatch = argRegex.exec(arg.trim())) !== null) {
      argList.push(argMatch[0]);
    }
  }

  const argValues: ArgValueType[] = [];
  const cellAddresses: Set<string> = new Set();
  let sheet: Excel.Worksheet | string | undefined;
  let cell: Excel.Range | undefined;
  let cellLoadedSuccessfully: boolean = true;

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
        let cellValue: string | number | Date = cell.values[0][0];

        if (
          cellNumberFormat &&
          cellNumberFormat.includes("yy") &&
          typeof cellValue === "number"
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

function groupCellsIntoRanges(cells: string[]): (string | null)[] {
  if (cells.length === 0) return [];

  const sortedCells = cells.sort((a: string, b: string) => {
    const colA = a.match(/[A-Z]+/)?.[0];
    const colB = b.match(/[A-Z]+/)?.[0];
    const rowA = parseInt(a.match(/\d+/)?.[0] ?? "0", 10);
    const rowB = parseInt(b.match(/\d+/)?.[0] ?? "0", 10);

    if (colA === colB) {
      return rowA - rowB;
    }

    return (colA ?? "").localeCompare(colB ?? "");
  });

  const ranges: string[] = [];
  let startCell = sortedCells[0];
  let endCell = sortedCells[0];

  for (let i = 1; i < sortedCells.length; i += 1) {
    const currentCell = sortedCells[i];
    const [currentColumn, currentRow] = [
      currentCell.match(/[A-Z]+/)?.[0],
      parseInt(currentCell.match(/\d+/)?.[0] ?? "0", 10),
    ];
    const [endColumn, endRow] = [
      endCell.match(/[A-Z]+/)?.[0],
      parseInt(endCell.match(/\d+/)?.[0] ?? "0", 10),
    ];

    if (
      (currentColumn === endColumn && currentRow === endRow + 1) ||
      (currentRow === endRow &&
        (currentColumn ?? "").charCodeAt(0) ===
          (endColumn ?? "").charCodeAt(0) + 1)
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
    const [startRange, endRange] = currentRange?.split(":") ?? [];
    const [nextStartRange, nextEndRange] = rangeCells[i].split(":");

    const currentStartCol = startRange.match(/[A-Z]+/)?.[0];
    const currentEndCol = endRange.match(/[A-Z]+/)?.[0];
    const nextStartCol = nextStartRange.match(/[A-Z]+/)?.[0];
    const nextEndCol = nextEndRange.match(/[A-Z]+/)?.[0];

    const currentStartRow = parseInt(startRange.match(/\d+/)?.[0] ?? "0", 10);
    const currentEndRow = parseInt(endRange.match(/\d+/)?.[0] ?? "0", 10);
    const nextStartRow = parseInt(nextStartRange.match(/\d+/)?.[0] ?? "0", 10);
    const nextEndRow = parseInt(nextEndRange.match(/\d+/)?.[0] ?? "0", 10);

    if (
      (currentEndCol ?? "").charCodeAt(0) + 1 ===
        (nextStartCol ?? "").charCodeAt(0) &&
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

function getArgs(formula: string, startIndex: number): string {
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

function splitArgs(args: string): string[] {
  const result: string[] = [];
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

function sortStepsByCalculationOrder(steps: GraphNodeType[]): GraphNodeType[] {
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

function applyFunctionSpecificLogic(
  funcName: string,
  args: string,
  formulaOrderInfo: GraphNodeType,
) {
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
  getArgs,
  splitArgs,
};
