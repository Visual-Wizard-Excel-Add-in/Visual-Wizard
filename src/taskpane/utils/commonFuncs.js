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

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "formulas", "values", "numberFormat"]);
      await context.sync();
async function loadCellInfo() {


      const selectCell = new CellInfo(range);

      updateCellState();

      function updateCellState() {
        const stateMapping = {
          cellAddress: { value: selectCell.address, setter: "setCellAddress" },
          cellValue: { value: selectCell.values, setter: "setCellValue" },
          cellFormula: { value: selectCell.formula, setter: "setCellFormula" },
    range.arguments = argumentsList();

        Object.keys(stateMapping).forEach((state) => {
          const { value, setter } = stateMapping[state];

          if (isChanged(value, usePublicStore.getState()[state])) {
            updateState(setter, value);
          }
        });

          return cellValue !== state;
        }
    function precedentsOrNull() {
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
    function argumentsList() {
      if (range.formulas[0][0]) {
        return precedents.addresses[0].split(",");
      }

    if (numberFormat && numberFormat.includes("yy") && targetCellValue !== "") {
      targetCellValue = new Date(
        (targetCellValue - 25569) * 86400 * 1000,
      ).toLocaleDateString();
      return [];
    }

    return targetCellValue;
  });

  return targetValue;
}

  let rangeValue = null;

  await Excel.run(async (context) => {
    const selectRange = context.workbook.getSelectedRange();

    selectRange.load("values");
    await context.sync();

    rangeValue = selectRange.values;
  });

  return rangeValue;
}

  const selectionHandler = useHandlerStore.getState().selectionChangeHandler;

  if (selectionHandler) {
    removeHandler(selectionHandler, "setSelectionChangeHandler");
async function updateCellInfo() {
    cellFunctions: {
      value: cellInfo.functions,
      setter: "setCellFunctions",
    },
    cellArgument: { value: cellInfo.arguments, setter: "setCellArguments" },
  };

  }

  await Excel.run(async (context) => {
    const { workbook } = context;
    const sheet = workbook.worksheets.getItem(sheetId);

    const handler = sheet.onSelectionChanged.add(func);
    await context.sync();

    useHandlerStore.getState().setSelectionChangeHandler(handler);
  });
}

  const match = address.match(/\$?([A-Z]+)\$?([0-9]+)/);

  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  function isStateChanged(cellValue, state) {
  }

  return [match[1], parseInt(match[2], 10)];
}

async function targetCellValue(targetCell) {



      }


  }

  function address() {
    if (sheetName()) {
      return targetCell.split("!")[1];
    }

  }


  }

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



}

async function addOnSelectionChange(sheetId, func) {



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
