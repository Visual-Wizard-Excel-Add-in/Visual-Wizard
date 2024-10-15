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
  getArgs,
  splitArgs,
};
