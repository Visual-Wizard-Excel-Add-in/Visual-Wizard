import { extractArgsAddress } from "./cellCommonUtils";

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

      await changeCellborder(resultCellRange, "red", false);
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

        await changeCellborder(argcellsRange, "red", false);
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

async function changeCellborder(targetCell, color, isClear) {
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

function convertAddressToA1(originalAddress, startRow, startColumn) {
  const row = parseInt(originalAddress.match(/\d+/)[0], 10) - startRow + 1;
  const column =
    originalAddress.match(/[A-Z]+/)[0].charCodeAt(0) - 65 - startColumn;
  const newAddress = String.fromCharCode(65 + column) + row.toString();

  return newAddress;
}

async function saveCellStylePreset(styleName) {
  try {
    await Excel.run(async (context) => {
      let cellStylePreset =
        Office.context.document.settings.get("cellStylePreset");

      if (!cellStylePreset) {
        cellStylePreset = {};
      } else {
        cellStylePreset = JSON.parse(cellStylePreset);
      }

      if (cellStylePreset[styleName]) {
        delete cellStylePreset[styleName];
      }

      const range = context.workbook.getSelectedRange();

      range.load(["rowCount", "columnCount", "address"]);
      await context.sync(); // 첫 번째 동기화

      const rows = range.rowCount;
      const columns = range.columnCount;
      const startAddress = range.address.split("!")[1];
      const startRow = parseInt(startAddress.match(/\d+/)[0], 10);
      const startColumn = startAddress.match(/[A-Z]+/)[0].charCodeAt(0) - 65;

      const cellStyles = {};

      for (let i = 0; i < rows; i += 1) {
        for (let j = 0; j < columns; j += 1) {
          const cell = range.getCell(i, j);

          cell.load("address");
          cell.format.load(["fill", "font", "borders", "protection"]);
          cell.format.fill.load("color");
          cell.format.font.load([
            "name",
            "size",
            "color",
            "bold",
            "italic",
            "underline",
            "strikethrough",
          ]);
          cell.format.protection.load(["locked", "formulaHidden"]);
          cell.load(["numberFormat"]);

          if (cell.format.alignment) {
            cell.format.alignment.load([
              "horizontal",
              "vertical",
              "wrapText",
              "indentLevel",
              "readingOrder",
            ]);
          }

          const borderEdges = [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
          ];
          const borders = {};

          borderEdges.forEach((edge) => {
            const border = cell.format.borders.getItem(edge);

            border.load(["style", "color", "weight"]);

            borders[edge] = border;
          });

          await context.sync();

          const originalAddress = cell.address.split("!")[1];
          const newAddress = convertAddressToA1(
            originalAddress,
            startRow,
            startColumn,
          );

          cellStyles[newAddress] = {
            font: {
              name: cell.format.font.name,
              size: cell.format.font.size,
              color: cell.format.font.color,
              bold: cell.format.font.bold,
              italic: cell.format.font.italic,
              underline: cell.format.font.underline,
              strikethrough: cell.format.font.strikethrough,
            },
            fill: {
              color: cell.format.fill.color,
            },
            alignment: cell.format.alignment
              ? {
                  horizontal: cell.format.alignment.horizontal,
                  vertical: cell.format.alignment.vertical,
                  wrapText: cell.format.alignment.wrapText,
                  indentLevel: cell.format.alignment.indentLevel,
                  readingOrder: cell.format.alignment.readingOrder,
                }
              : {},
            numberFormat: cell.numberFormat,
            borders: {},
            protection: {
              locked: cell.format.protection.locked,
              formulaHidden: cell.format.protection.formulaHidden,
            },
          };

          borderEdges.forEach((edge) => {
            const border = borders[edge];

            if (border.style === Excel.BorderLineStyle.none) {
              cellStyles[newAddress].borders[
                edge.replace("Edge", "").toLowerCase()
              ] = {
                style: border.style,
                color: "#d6d6d6",
                weight: "thin",
              };
            } else {
              cellStyles[newAddress].borders[
                edge.replace("Edge", "").toLowerCase()
              ] = {
                style: border.style,
                color: border.color,
                weight: border.weight,
              };
            }
          });
        }
      }

      cellStylePreset[styleName] = cellStyles;

      Office.context.document.settings.set(
        "cellStylePreset",
        JSON.stringify(cellStylePreset),
      );
      await Office.context.document.settings.saveAsync();
    });
  } catch (error) {
    throw new Error("Error in saveCellStylePreset:", error);
  }
}

async function loadCellStylePreset(styleName) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "rowCount", "columnCount"]);
      await context.sync();

      let cellStylePreset =
        Office.context.document.settings.get("cellStylePreset");

      if (!cellStylePreset) {
        return;
      }

      cellStylePreset = JSON.parse(cellStylePreset);
      const savedCellStyles = cellStylePreset[styleName];

      if (!savedCellStyles) {
        return;
      }

      const rows = range.rowCount;
      const columns = range.columnCount;
      const savedAddresses = Object.keys(savedCellStyles);
      const savedRows = savedAddresses.reduce((max, addr) => {
        const row = parseInt(addr.match(/\d+$/)[0], 10);
        return Math.max(max, row);
      }, 0);
      const savedColumns = savedAddresses.reduce((max, addr) => {
        const col = addr.match(/^[A-Z]+/)[0];
        return Math.max(max, col.charCodeAt(0) - 64);
      }, 0);
      const getAddress = (row, column) => {
        const col = String.fromCharCode(64 + column);
        return `${col}${row}`;
      };
      const getStyle = (row, column) => {
        const address = getAddress(row, column);
        return savedCellStyles[address] || {};
      };
      const applyStyle = (cell, styles) => {
        if (styles.font) {
          Object.keys(styles.font).forEach((key) => {
            if (styles.font[key] !== undefined) {
              cell.format.font[key] = styles.font[key];
            }
          });
        }

        if (styles.fill && styles.fill.color) {
          cell.format.fill.color = styles.fill.color;
        }

        if (styles.borders) {
          ["top", "bottom", "left", "right"].forEach((edge) => {
            if (styles.borders[edge]) {
              const border = cell.format.borders.getItem(
                `Edge${edge.charAt(0).toUpperCase() + edge.slice(1)}`,
              );

              if (styles.borders[edge].style) {
                border.style = styles.borders[edge].style;
              }

              if (styles.borders[edge].color) {
                border.color = styles.borders[edge].color;
              }

              if (styles.borders[edge].weight) {
                border.weight = styles.borders[edge].weight;
              }
            }
          });
        }

        if (styles.alignment) {
          const alignmentProperties = [
            "horizontal",
            "vertical",
            "wrapText",
            "shrinkToFit",
            "indentLevel",
            "readingOrder",
          ];

          alignmentProperties.forEach((prop) => {
            if (styles.alignment[prop] !== undefined) {
              cell.format.alignment[prop] = styles.alignment[prop];
            }
          });
        }

        if (styles.protection) {
          if (styles.protection.locked !== undefined) {
            cell.format.protection.locked = styles.protection.locked;
          }

          if (styles.protection.formulaHidden !== undefined) {
            cell.format.protection.formulaHidden =
              styles.protection.formulaHidden;
          }
        }

        if (styles.numberFormat) {
          cell.numberFormat = styles.numberFormat;
        }
      };

      for (let i = 1; i <= rows; i += 1) {
        for (let j = 1; j <= columns; j += 1) {
          const cell = range.getCell(i - 1, j - 1);
          const rowInSavedRange = ((i - 1) % savedRows) + 1;
          const columnInSavedRange = ((j - 1) % savedColumns) + 1;
          const styles = getStyle(rowInSavedRange, columnInSavedRange);
          applyStyle(cell, styles);
        }
      }

      await context.sync();
    });
  } catch (error) {
    throw new Error(
      "Error in loadCellStylePreset:",
      error.message,
      error.stack,
    );
  }
}

export {
  storeCellStyle,
  applyCellStyle,
  highlightingCell,
  changeCellborder,
  saveCellStylePreset,
  loadCellStylePreset,
};
