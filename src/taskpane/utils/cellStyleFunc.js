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

export { storeCellStyle, applyCellStyle, highlightingCell, ChangeCellborder };
