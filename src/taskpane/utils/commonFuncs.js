import usePublicStore from "../store/publicStore";
import useHandlerStore from "../store/handlerStore";
import CellInfo from "../classes/CellInfo";
import MESSAGE_LIST from "../constants/messageConstants";

function updateState(setStateFunc, newValue) {
  usePublicStore.getState()[setStateFunc](newValue);
}

async function removeHandler(handler, setter) {
  try {
    if (handler) {
      await Excel.run(handler.context, async (context) => {
        handler.remove();
        await context.sync();
      });

      useHandlerStore.getState()[setter](null);
    }
  } catch (error) {
    throw new Error(error);
  }
}

async function loadCellInfo() {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const isNull = range.getUsedRangeOrNullObject();

    range.load(["address", "formulas", "values", "numberFormat"]);
    await context.sync();

    const precedents = await precedentsOrNull(range, isNull);
    await context.sync();

    range.arguments = argumentsList(range, precedents);

    return new CellInfo(range);
  });
}

async function precedentsOrNull(range, isNull) {
  let result = null;

  if (isNull.isNullObject) {
    return result;
  }

  result = range.getDirectPrecedents();

  if (result && range.formulas[0][0]) {
    try {
      result.load("addresses");
    } catch (e) {
      return result;
    }
  }

  return result;
}

function argumentsList(range, precedents) {
  const formula = range.formulas[0][0];

  if (precedents && formula) {
    return precedents.addresses[0]
      .split(",")
      .map((arg) => arg.replaceAll("'", ""));
  }

  if (formula.includes("(")) {
    return range.formulas[0][0].split("(")[1].split(")")[0].split(",").trim();
  }

  return [];
}

async function updateCellInfo() {
  const cellInfo = await loadCellInfo();

  const stateMapping = {
    cellAddress: { value: cellInfo.address, setter: "setCellAddress" },
    cellValue: { value: cellInfo.values, setter: "setCellValue" },
    cellFormula: { value: cellInfo.formula, setter: "setCellFormula" },
    cellFunctions: {
      value: cellInfo.functions,
      setter: "setCellFunctions",
    },
    cellArgument: { value: cellInfo.arguments, setter: "setCellArguments" },
  };

  Object.keys(stateMapping).forEach((state) => {
    const { value, setter } = stateMapping[state];

    if (isStateChanged(value, state)) {
      updateState(setter, value);
    }
  });

  function isStateChanged(cellValue, state) {
    return cellValue !== usePublicStore.getState()[state];
  }
}

async function targetCellValue(targetCell) {
  try {
    return await Excel.run(async (context) => {
      const cell = sheet().getRange(address());

      cell.load(["values", "numberFormat"]);
      await context.sync();

      let result = "";

      if (isDateFormat(cell)) {
        result = dateValue(cell);
      } else {
        [[result]] = cell.values;
      }

      return result;

      function sheet() {
        return sheetName()
          ? context.workbook.worksheets.getItem(sheetName())
          : context.workbook.worksheets.getActiveWorksheet();
      }
    });
  } catch (error) {
    throw new Error(error);
  }

  function address() {
    if (sheetName()) {
      return targetCell.split("!")[1];
    }

    return targetCell.split("!")[0];
  }

  function sheetName() {
    if (targetCell.split("!").length > 1) {
      return targetCell.split("!")[0];
    }

    return null;
  }

  function dateValue(cell) {
    return new Intl.DateTimeFormat("ko-KR").format(
      (cell.values[0][0] - 25569) * 86400 * 1000,
    );
  }

  function isDateFormat(cell) {
    const [[numberFormat]] = cell.numberFormat;

    return (
      numberFormat?.includes("yy") ||
      numberFormat.includes("dd") ||
      numberFormat.includes("mm")
    );
  }
}

async function selectRangeValues() {
  return await Excel.run(async (context) => {
    let result = null;
    const selectRange = context.workbook.getSelectedRange();

    selectRange.load("values");
    await context.sync();

    result = selectRange.values;

    return result;
  });
}

async function addOnSelectionChange(sheetId, func) {
  const existHandler = useHandlerStore.getState().selectionChangeHandler;

  if (existHandler) {
    removeHandler(existHandler, "setSelectionChangeHandler");
  }

  await Excel.run(async (context) => {
    const { workbook } = context;
    const sheet = workbook.worksheets.getItem(sheetId);

    const newHandler = sheet.onSelectionChanged.add(func);
    await context.sync();

    useHandlerStore.getState().setSelectionChangeHandler(newHandler);
  });
}

function popUpMessage(purpose = "default", option = "") {
  const message = MESSAGE_LIST[purpose];

  message.body += `\n${option}`;

  updateState("setMessageList", message);
}

export {
  addOnSelectionChange,
  updateCellInfo,
  selectRangeValues,
  updateState,
  targetCellValue,
  popUpMessage,
  removeHandler,
};
