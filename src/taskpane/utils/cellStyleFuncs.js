import { popUpMessage } from "./commonFuncs";
import STYLE_OPTIONS_TO_LOAD from "../constants/styleConstants";

async function extractCellStyle(context, address) {
  try {
    const targetRange = await determineTarget();
    const properties = targetRange.getCellProperties(STYLE_OPTIONS_TO_LOAD);

    await context.sync();

    const styleProperties = properties.value.map((row) =>
      row.map((cell) => getMainBorders(cell)),
    );

    await context.sync();

    return styleProperties;
  } catch (error) {
    popUpMessage("workFail", error.message);

    throw new Error(error.message);
  }

  async function determineTarget() {
    if (typeof address === "string") {
      const target = context.workbook.worksheets.getRanges(address);

      await context.sync();

      return target;
    }

    return address;
  }

  function getMainBorders(cell) {
    const { format } = cell;
    const { borders } = format;
    const mainBorders = ["bottom", "top", "left", "right"];
    const filteredBorders = {};

    for (const border in borders) {
      if (mainBorders.includes(border) && borders[border].style === "None") {
        filteredBorders[border] = borders[border];
        filteredBorders[border].color = "#D6D6D6";
        filteredBorders[border].style = "Continuous";
        filteredBorders[border].tintAndShade = 0;
      } else if (
        mainBorders.includes(border) ||
        borders[border].style !== "None"
      ) {
        filteredBorders[border] = borders[border];
      }
    }

    return {
      format: {
        ...format,
        borders: filteredBorders,
      },
    };
  }
}

async function storeCellStyle(cellAddress, allPresets, isHighlight) {
  let cellStyleToReturn = null;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);

      const allSavedPresets = await OfficeRuntime.storage.getItem(allPresets);
      const parsedPresets = allSavedPresets
        ? { ...JSON.parse(allSavedPresets) }
        : {};

      const cellStyle = await extractCellStyle(context, cell);

      if (parsedPresets[cellAddress] && !isHighlight) {
        return;
      }

      if (!parsedPresets[cellAddress] && isHighlight) {
        parsedPresets[cellAddress] = cellStyle;
      }

      if (allPresets === "allMacroPresets") {
        cellStyleToReturn = cellStyle;
      } else {
        await OfficeRuntime.storage.setItem(
          allPresets,
          JSON.stringify(parsedPresets),
        );
      }
    });

    if (allPresets === "allMacroPresets") {
      return cellStyleToReturn;
    }
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error.message);
  }
  return null;
}

async function applyCellStyle(
  cellAddress,
  allPresets,
  isHighlight,
  actionCellStyle = null,
) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);
      const allSavedPresets = await OfficeRuntime.storage.getItem(allPresets);
      let cellStyle = {};

      if (!allSavedPresets) {
        popUpMessage("loadFail", "저장된 프리셋이 없습니다!");

        throw new Error("No any saved presets found.");
      }

      const parsedPresets = JSON.parse(allSavedPresets);

      if (allPresets === "allMacroPresets") {
        cellStyle = actionCellStyle;
      } else {
        cellStyle = parsedPresets[cellAddress];
      }

      if (cellStyle && !isHighlight) {
        cell.setCellProperties(cellStyle);

        await context.sync();

        delete parsedPresets[cellAddress];

        await OfficeRuntime.storage.setItem(
          allPresets,
          JSON.stringify(parsedPresets),
        );
      }
    });
  } catch (e) {
    popUpMessage("loadFail", e.message);

    throw new Error(e.message);
  }
}

async function detectErrorCell(isHighlight) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      const errorTypes = Object.values(Excel.ErrorCellValueType).map((error) =>
        error.toUpperCase(),
      );

      range.load("values, address");
      await context.sync();

      const errorCells = [];

      if (range.values) {
        range.values.forEach((row, rowIndex) => {
          row.forEach((cell, colIndex) => {
            const parsedCellValue =
              typeof cell === "string" &&
              cell?.split("#")[1]?.split("!")[0].split("/").join("")
                ? cell?.split("#")[1]?.split("!")[0].split("/").join("")
                : null;

            if (errorTypes.includes(parsedCellValue)) {
              const cellRange = range.getCell(rowIndex, colIndex);

              errorCells.push(cellRange);
            }
          });
        });
      }

      errorCells.forEach(async (cell) => cell.load("address"));
      await context.sync();

      for (const cell of errorCells) {
        const cellAddress = cell.address;

        if (isHighlight) {
          await storeCellStyle(cellAddress, "allCellStyles", true);
        } else {
          await applyCellStyle(cellAddress, "allCellStyles", false);
        }

        await context.sync();
      }

      for (const cell of errorCells) {
        if (isHighlight) {
          cell.format.fill.color = "red";

          const edges = ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"];

          edges.forEach((edge) => {
            const border = cell.format.borders.getItem(edge);
            border.color = "green";
            border.style = Excel.BorderLineStyle.continuous;
            border.weight = Excel.BorderWeight.thick;
          });
        }

        await context.sync();
      }
    });
  } catch (error) {
    popUpMessage("workFail", error.message);

    throw new Error(error.message);
  }
}

async function highlightingCell(isHighlight, resultCell) {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectRange = context.workbook.getSelectedRange();
    const getPrecedents = selectRange.getDirectPrecedents();
    const argsAddress = [];

    selectRange.load("address");
    getPrecedents.areas.load("address");

    await context.sync();

    for (let i = 0; i < getPrecedents.areas.items.length; i += 1) {
      if (getPrecedents.areas.items[i].address.includes(",")) {
        getPrecedents.areas.items[i].address
          .split(",")
          .forEach((address) => argsAddress.push(address));
      } else {
        argsAddress.push(getPrecedents.areas.items[i].address);
      }
    }

    const resultFill = {
      fill: {
        color: "#3d33ff",
      },
    };
    const argsFill = {
      fill: {
        color: "#28f925",
      },
    };
    const highlightBorder = {
      borders: {
        bottom: {
          color: "red",
          weight: Excel.BorderWeight.thick,
          style: Excel.BorderLineStyle.continuous,
        },
        top: {
          color: "red",
          weight: Excel.BorderWeight.thick,
          style: Excel.BorderLineStyle.continuous,
        },
        left: {
          color: "red",
          weight: Excel.BorderWeight.thick,
          style: Excel.BorderLineStyle.continuous,
        },
        right: {
          color: "red",
          weight: Excel.BorderWeight.thick,
          style: Excel.BorderLineStyle.continuous,
        },
      },
    };

    if (isHighlight) {
      await storeCellStyle(resultCell, "allCellStyles", isHighlight);

      for (let i = 0; i < argsAddress.length; i += 1) {
        await storeCellStyle(argsAddress[i], "allCellStyles", isHighlight);
      }

      const rangesToLoad = argsAddress.map((address) => {
        const targetRange = worksheet.getRange(address);

        targetRange.load("values");

        return targetRange;
      });

      await context.sync();

      const requests = rangesToLoad.map(async (targetRange) => {
        const argsStyle = targetRange.values;

        const argsHighilighStyle = argsStyle.map((row) =>
          row.map(() => ({ format: { ...argsFill, ...highlightBorder } })),
        );

        return targetRange.setCellProperties(argsHighilighStyle);
      });

      await Promise.allSettled(requests);

      selectRange.setCellProperties([
        [{ format: { ...resultFill, ...highlightBorder } }],
      ]);
    } else {
      await applyCellStyle(resultCell, "allCellStyles", isHighlight);

      const requests = argsAddress.map(async (targetRange) => {
        await applyCellStyle(targetRange, "allCellStyles", isHighlight);
      });

      await Promise.allSettled(requests);
    }

    await context.sync();
  });
}

async function copyRangeStyle(presetName) {
  try {
    if (!presetName) {
      popUpMessage("saveFail", "프리셋을 정확히 선택해주세요!");

      return;
    }

    await Excel.run(async (context) => {
      let cellStylePresets =
        await OfficeRuntime.storage.getItem("cellStylePresets");

      if (!cellStylePresets) {
        cellStylePresets = {};
      } else {
        cellStylePresets = JSON.parse(cellStylePresets);
      }

      if (cellStylePresets[presetName]) {
        delete cellStylePresets[presetName];
      }

      const range = context.workbook.getSelectedRange();

      range.load("address");
      await context.sync();

      const cellAddress = range.address.split("!")[1];
      const cellStyles = await extractCellStyle(context, range);

      cellStylePresets[presetName] = [cellStyles, cellAddress];

      await OfficeRuntime.storage.setItem(
        "cellStylePresets",
        JSON.stringify(cellStylePresets),
      );

      popUpMessage("saveSuccess");
    });
  } catch (error) {
    popUpMessage("saveFail", error.message);

    throw new Error(error.message);
  }
}

async function pasteRangeStyle(styleName) {
  try {
    await Excel.run(async (context) => {
      const selectRange = context.workbook.getSelectedRange();
      let styleSheet =
        context.workbook.worksheets.getItemOrNullObject("StyleSheet");

      await context.sync();

      let cellStylePresets =
        await OfficeRuntime.storage.getItem("cellStylePresets");

      cellStylePresets = JSON.parse(cellStylePresets);
      const savedCellStyles = cellStylePresets[styleName][0];
      const savedCellAddress = cellStylePresets[styleName][1];

      if (!savedCellStyles) {
        popUpMessage("loadFail", "저장된 서식이 없습니다!");

        throw new Error("Preset not found");
      }

      if (!styleSheet.isNullObject) {
        styleSheet.delete();
        await context.sync();
      }

      selectRange.load(["address"]);
      styleSheet = context.workbook.worksheets.add("StyleSheet");
      await context.sync();

      const sourceRange = styleSheet.getRange(savedCellAddress);

      sourceRange.setCellProperties(savedCellStyles);

      selectRange.copyFrom(sourceRange, "Formats");

      styleSheet.delete();
      await context.sync();
    });
  } catch (error) {
    popUpMessage("loadFail", error.message);

    throw new Error(error.message);
  }
}

export {
  storeCellStyle,
  applyCellStyle,
  highlightingCell,
  copyRangeStyle,
  pasteRangeStyle,
  detectErrorCell,
};
