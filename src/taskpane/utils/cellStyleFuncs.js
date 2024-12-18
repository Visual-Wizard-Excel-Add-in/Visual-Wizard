import useTotalStore from "../store/useTotalStore";
import { popUpMessage } from "./commonFuncs";
import {
  STYLE_OPTIONS_TO_LOAD,
  HIGHLIGHT_STYLES,
} from "../constants/styleConstants";

async function extractCellStyle(context, rangeObject) {
  try {
    const rawProperties = rangeObject.getCellProperties(STYLE_OPTIONS_TO_LOAD);

    await context.sync();

    const result = rawProperties.value.map((row) =>
      row.map((cell) => filterBorders(cell)),
    );

    return result;
  } catch (error) {
    popUpMessage("workFail", error.message);

    throw new Error(error.message);
  }

  function filterBorders(cell) {
    const { format } = cell;
    const mainBorders = ["bottom", "top", "left", "right"];
    const filteredBorders = {};

    Object.keys(format.borders).map((border) => nomalizeEmptyBorders(border));

    const result = {
      format: {
        ...format,
        borders: filteredBorders,
      },
    };

    return result;

    function nomalizeEmptyBorders(border) {
      if (isEmptyMainBorder()) {
        filteredBorders[border] = { ...format.borders[border] };

        const mainBorder = filteredBorders[border];

        mainBorder.color = "#D6D6D6";
        mainBorder.style = "Continuous";
        mainBorder.tintAndShade = 0;
      } else if (isStyledBorder()) {
        filteredBorders[border] = format.borders[border];
      }

      function isEmptyMainBorder() {
        return (
          mainBorders.includes(border) &&
          format.borders[border].style === "None"
        );
      }

      function isStyledBorder() {
        return (
          mainBorders.includes(border) ||
          format.borders[border].style !== "None"
        );
      }
    }
  }
}

async function storeCellStyle(address, PresetType, isHighlight) {
  if (!isHighlight) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(address);
      const loadedPresets = await loadPreset();

      loadedPresets[address] = await extractCellStyle(context, cell);

      await OfficeRuntime.storage.setItem(
        PresetType,
        JSON.stringify(loadedPresets),
      );
    });
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error.message);
  }

  async function loadPreset() {
    const existingData = JSON.parse(
      await OfficeRuntime.storage.getItem(PresetType),
    );

    if (existingData) {
      const result = { ...existingData };

      return result;
    }

    return {};
  }
}

async function restoreCellStyle(
  address,
  presetType,
  isHighlight,
  actionCellStyle = null,
) {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(address);
      const cellStyle = await loadCellStyle();

      await applyStyle(cell, cellStyle, context);
    } catch (e) {
      popUpMessage("loadFail", e.message);

      throw new Error(e.message);
    }
  });

  async function applyStyle(cell, cellStyle, context) {
    if (cellStyle && !isHighlight && actionCellStyle) {
      cell.setCellProperties(cellStyle);
      await context.sync();
    } else {
      cell.setCellProperties(cellStyle[address]);
      await context.sync();

      delete cellStyle[address];

      await OfficeRuntime.storage.setItem(
        presetType,
        JSON.stringify(cellStyle),
      );
    }
  }

  async function loadCellStyle() {
    let result = {};

    if (actionCellStyle) {
      result = actionCellStyle;
    } else {
      const loadedPresets = await OfficeRuntime.storage.getItem(presetType);

      if (!loadedPresets) {
        popUpMessage("loadFail", "저장된 프리셋이 없습니다!");
      }

      result = JSON.parse(loadedPresets);
    }

    return result;
  }
}

async function detectErrorCell(isHighlight) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();

      range.load("values");
      await context.sync();

      const errorCells = findError(range, errorList());

      errorCells.forEach((cell) => cell.load("address"));
      await context.sync();

      for (const cell of errorCells) {
        const cellData = [cell.address, "allCellStyles", isHighlight];

        if (isHighlight) {
          await storeCellStyle(...cellData);
        } else {
          await restoreCellStyle(...cellData);
        }
      }

      if (isHighlight) {
        highlightError(errorCells);
      }

      await context.sync();
    });
  } catch (error) {
    popUpMessage("workFail", error.message);

    throw new Error(error.message);
  }

  function errorType(cell) {
    return cell?.split("#")[1]?.split("!")[0].split("/").join("");
  }

  function errorList() {
    return Object.values(Excel.ErrorCellValueType).map((error) =>
      error.toUpperCase(),
    );
  }

  function findError(range, errorTypes) {
    const result = [];

    if (range.values) {
      range.values.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const cellValue =
            typeof cell === "string" && errorType(cell)
              ? errorType(cell)
              : null;

          if (errorTypes.includes(cellValue)) {
            const cellRange = range.getCell(rowIndex, colIndex);

            result.push(cellRange);
          }
        });
      });
    }

    return result;
  }

  function highlightError(errorCells) {
    errorCells.forEach((cell) => {
      cell.format.fill.color = "red";

      const edges = ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"];

      edges.forEach((edge) => {
        const border = cell.format.borders.getItem(edge);

        border.color = "green";
        border.style = Excel.BorderLineStyle.continuous;
        border.weight = Excel.BorderWeight.thick;
      });
    });
  }
}

async function highlightingCell(isHighlight, resultCell) {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectRange = context.workbook.getSelectedRange();
    const argsAddress = useTotalStore.getState().cellArguments;

    selectRange.load("address");

    await context.sync();

    if (isHighlight) {
      await storeCellStyle(resultCell, "allCellStyles", isHighlight);

      for (const address of argsAddress) {
        await storeCellStyle(address, "allCellStyles", isHighlight);
      }

      await highlightArgCells();
      highlightResultCell();
    } else {
      await restoreCellStyle(resultCell, "allCellStyles", isHighlight);

      for (const address of argsAddress) {
        await restoreCellStyle(address, "allCellStyles", isHighlight);
      }
    }

    await context.sync();

    function highlightResultCell() {
      selectRange.setCellProperties([
        [
          {
            format: {
              ...HIGHLIGHT_STYLES.resultFill,
              borders: HIGHLIGHT_STYLES.borders,
            },
          },
        ],
      ]);
    }

    async function highlightArgCells() {
      const rangesToLoad = argsAddress.map((address) => {
        const targetRange = worksheet.getRange(address);

        targetRange.load("values");

        return targetRange;
      });

      await context.sync();

      await Promise.allSettled(
        rangesToLoad.map(async (targetRange) => {
          const argsHighilighStyle = targetRange.values.map((row) =>
            row.map(() => ({
              format: {
                ...HIGHLIGHT_STYLES.argsFill,
                borders: HIGHLIGHT_STYLES.borders,
              },
            })),
          );

          return targetRange.setCellProperties(argsHighilighStyle);
        }),
      );
    }
  });
}

async function copyRangeStyle(presetName) {
  try {
    if (!presetName) {
      popUpMessage("saveFail", "프리셋을 정확히 선택해주세요!");

      return;
    }

    await Excel.run(async (context) => {
      const styleList = loadPresets();

      const range = context.workbook.getSelectedRange();

      range.load("address");
      await context.sync();

      const cellAddress = range.address.split("!")[1];
      const cellStyles = await extractCellStyle(context, range);

      styleList[presetName] = [cellStyles, cellAddress];

      await OfficeRuntime.storage.setItem(
        "cellStylePresets",
        JSON.stringify(styleList),
      );

      popUpMessage("saveSuccess");
    });
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error.message);
  }

  async function loadPresets() {
    let result = JSON.parse(
      await OfficeRuntime.storage.getItem("cellStylePresets"),
    );

    if (result) {
      delete result[presetName];
    } else {
      result = {};
    }

    return result;
  }
}

async function pasteRangeStyle(presetName) {
  const stylePresets = JSON.parse(
    await OfficeRuntime.storage.getItem("cellStylePresets"),
  );

  try {
    await Excel.run(async (context) => {
      await checkDuplicate();

      const [savedStyles, savedAddress] = stylePresets[presetName];

      if (!savedStyles) {
        popUpMessage("loadFail", "저장된 서식이 없습니다!");

        throw new Error("Preset not found");
      }

      const selectRange = context.workbook.getSelectedRange();
      const sourceSheet = context.workbook.worksheets.add("StyleSheet");
      const sourceRange = sourceSheet.getRange(savedAddress);

      sourceRange.setCellProperties(savedStyles);
      selectRange.copyFrom(sourceRange, "Formats");
      sourceSheet.delete();
      await context.sync();

      async function checkDuplicate() {
        const source =
          context.workbook.worksheets.getItemOrNullObject("StyleSheet");

        await context.sync();

        if (!source.isNullObject) {
          source.delete();
          await context.sync();
        }
        return source;
      }
    });
  } catch (error) {
    popUpMessage("loadFail", error.message);

    throw new Error(error.message);
  }
}

export {
  storeCellStyle,
  restoreCellStyle,
  highlightingCell,
  copyRangeStyle,
  pasteRangeStyle,
  detectErrorCell,
  extractCellStyle,
};
