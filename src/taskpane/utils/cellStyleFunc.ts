import { extractArgsAddress, updateState } from "./cellCommonUtils";

async function storeCellStyle(
  cellAddress: string,
  allPresets: string,
  isCellHighlighting: Boolean,
) {
  let cellStyleToReturn: CellStyleType | null = null;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);
      const edges: EdgesType = [
        "EdgeBottom",
        "EdgeLeft",
        "EdgeTop",
        "EdgeRight",
      ];

      const allSavedPresets: string =
        await OfficeRuntime.storage.getItem(allPresets);
      const parsedPresets: StylePresetsType = allSavedPresets
        ? { ...JSON.parse(allSavedPresets) }
        : {};

      cell.load(["address", "numberFormat", "numberFormatLocal"]);
      cell.format.load([
        "fill/color",
        "font",
        "borders",
        "protection",
        "horizontalAlignment",
        "verticalAlignment",
        "wrapText",
        "indentLevel",
        "readingOrder",
        "textOrientation",
      ]);
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
      await context.sync();

      if (parsedPresets[cellAddress] && !isCellHighlighting) {
        return;
      }

      const borders: { [key: string]: Excel.RangeBorder } = {};

      for (const edge of edges) {
        borders[edge] = cell.format.borders.getItem(edge);

        borders[edge].load(["color", "style", "weight"]);
      }

      await context.sync();

      const cellStyle: CellStyleType = {
        borders: {},
        numberFormat: cell.numberFormat[0][0],
        numberFormatLocal: cell.numberFormatLocal[0][0],
        font: {
          name: cell.format.font.name,
          bold: cell.format.font.bold,
          color: cell.format.font.color,
          italic: cell.format.font.italic,
          size: cell.format.font.size,
          underline: cell.format.font.underline,
          tintAndShade: cell.format.font.tintAndShade,
        },
        fill: {
          color: cell.format.fill.color,
        },
        alignment: {
          horizontalAlignment: cell.format.horizontalAlignment,
          verticalAlignment: cell.format.verticalAlignment,
          wrapText: cell.format.wrapText,
          indentLevel: cell.format.indentLevel,
          readingOrder: cell.format.readingOrder,
          textOrientation: cell.format.textOrientation,
        },
        protection: {
          locked: cell.format.protection.locked,
          formulaHidden: cell.format.protection.formulaHidden,
        },
      };

      for (const edge of edges) {
        const border = borders[edge];

        cellStyle.borders[edge] = {
          color: border.color,
          style: border.style,
          weight: border.weight,
        };
      }

      if (!parsedPresets[cellAddress] && isCellHighlighting) {
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
  } catch (e: unknown) {
    if (e instanceof Error) {
      const warningMessage = {
        type: "warning",
        title: "저장 실패: ",
        body: `기존 셀 서식을 저장하는데 실패했습니다. ${e.message}`,
      };

      updateState("setMessageList", warningMessage);
      throw new Error(e.message);
    }
  }
  return null;
}

async function applyCellStyle(
  cellAddress: string,
  allPresets: string,
  isCellHighlighting: boolean,
  actionCellStyle: CellStyleType | null = null,
) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);
      const edges: EdgesType = [
        "EdgeBottom",
        "EdgeLeft",
        "EdgeTop",
        "EdgeRight",
      ];
      const allSavedPresets: string =
        await OfficeRuntime.storage.getItem(allPresets);
      let cellStyle: CellStyleType | null = null;

      if (!allSavedPresets) {
        throw new Error("No any saved presets found.");
      }

      const parsedPresets: StylePresetsType = JSON.parse(allSavedPresets);

      if (allPresets === "allMacroPresets") {
        cellStyle = actionCellStyle;
      } else {
        cellStyle = parsedPresets[cellAddress];
      }

      if (cellStyle && !isCellHighlighting) {
        if (cellStyle.fill.color) {
          cell.format.fill.color = cellStyle.fill.color;
        }

        if (cellStyle.font.bold) {
          cell.format.font.bold = cellStyle.font.bold;
        }

        if (cellStyle.font.color) {
          cell.format.font.color = cellStyle.font.color;
        }

        if (cellStyle.font.tintAndShade) {
          cell.format.font.tintAndShade = cellStyle.font.tintAndShade;
        }
        cell.numberFormat[0][0] = cellStyle.numberFormat;
        cell.format.font.size = cellStyle.font.size;
        cell.format.font.underline = cellStyle.font.underline;
        cell.format.horizontalAlignment =
          cellStyle.alignment.horizontalAlignment;
        cell.format.verticalAlignment = cellStyle.alignment.verticalAlignment;

        if (cellStyle.borders) {
          for (const edge of edges) {
            if (Object.prototype.hasOwnProperty.call(cellStyle.borders, edge)) {
              const border = cell.format.borders.getItem(edge);
              const borderStyle = cellStyle.borders[edge];

              if (borderStyle.style !== "None") {
                border.color = borderStyle.color;
                border.style = borderStyle.style;
                border.weight = borderStyle.weight;
              } else {
                const infoMessage = {
                  type: "info",
                  title: "빈 테두리 서식: ",
                  body: "빈 테두리 서식은 기본 회색 테두리로 복원됩니다.",
                };

                updateState("setMessageList", infoMessage);

                border.style = Excel.BorderLineStyle.none;
                border.color = "#d6d6d6";
                border.weight = "Thin";
              }
            }
          }
        }

        await context.sync();

        if (allPresets !== "allMacroPresets") {
          delete parsedPresets[cellAddress];

          await OfficeRuntime.storage.setItem(
            allPresets,
            JSON.stringify(parsedPresets),
          );
        }
      }
    });
  } catch (e) {
    if (e instanceof Error) throw new Error(e.message);
  }
}

async function detectErrorCell(isCellHighlighting: boolean) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();

    range.load("values, address");
    await context.sync();

    const errorCells: Excel.Range[] = [];

    if (range.values) {
      range.values.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (
            cell === "#DIV/0!" ||
            cell === "#N/A" ||
            cell === "#VALUE!" ||
            cell === "#REF!" ||
            cell === "#NAME?" ||
            cell === "#NUM!" ||
            cell === "#NULL!"
          ) {
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

      if (isCellHighlighting) {
        await storeCellStyle(cellAddress, "allCellStyles", true);
      } else {
        await applyCellStyle(cellAddress, "allCellStyles", false);
      }

      await context.sync();
    }

    for (const cell of errorCells) {
      if (isCellHighlighting) {
        cell.format.fill.color = "red";

        const edges: EdgesType = [
          "EdgeBottom",
          "EdgeLeft",
          "EdgeTop",
          "EdgeRight",
        ];

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
}

async function highlightingCell(
  isCellHighlighting: boolean,
  argCells: string[],
  resultCell: string,
): Promise<void> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const resultCellRange = worksheet.getRange(resultCell);

    await storeCellStyle(resultCell, "allCellStyles", isCellHighlighting);

    const argsCellAddresses = argCells.map(extractArgsAddress).filter(Boolean);

    for (const argCell of argsCellAddresses) {
      if (typeof argCell === "string") {
        await storeCellStyle(argCell, "allCellStyles", isCellHighlighting);
      }
    }

    if (isCellHighlighting) {
      resultCellRange.format.fill.color = "#3d33ff";

      await changeCellborder(resultCellRange, "red", false);
    } else {
      await applyCellStyle(resultCell, "allCellStyles", isCellHighlighting);
    }

    for (const argCell of argsCellAddresses) {
      if (typeof argCell === "string") {
        const argcellsRange = worksheet.getRange(argCell);

        if (isCellHighlighting) {
          argcellsRange.format.fill.color = "#28f925";

          await changeCellborder(argcellsRange, "red", false);
        } else {
          await applyCellStyle(argCell, "allCellStyles", isCellHighlighting);
        }
      }
    }

    await context.sync();
  });
}

async function changeCellborder(
  targetCell: Excel.Range,
  color: string,
  isClear: boolean,
): Promise<void> {
  const edges: EdgesType = ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"];

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

function convertAddressToA1(
  originalAddress: string,
  startRow: number,
  startColumn: number,
): string {
  const row =
    parseInt(originalAddress.match(/\d+/)?.[0] ?? "0", 10) - startRow + 1;
  const columnMatch = originalAddress.match(/[A-Z]+/)?.[0];
  const column = columnMatch
    ? columnMatch.charCodeAt(0) - 65 - startColumn
    : -startColumn;
  const newAddress = String.fromCharCode(65 + column) + row.toString();

  return newAddress;
}

async function saveCellStylePreset(styleName: string) {
  try {
    if (styleName === "") {
      const warningMessage = {
        type: "warning",
        title: "접근 오류",
        body: "프리셋을 정확하게 선택해주세요",
      };

      updateState("setMessageList", warningMessage);

      return;
    }

    await Excel.run(async (context) => {
      let cellStylePresets: string =
        await OfficeRuntime.storage.getItem("cellStylePresets");

      const allStylePresets: { [key: string]: StylePresetsType } =
        JSON.parse(cellStylePresets);

      if (allStylePresets[styleName]) {
        delete allStylePresets[styleName];
      }

      const range = context.workbook.getSelectedRange();

      range.load(["rowCount", "columnCount", "address"]);
      await context.sync();

      const rows = range.rowCount;
      const columns = range.columnCount;
      const startAddress = range.address.split("!")[1];
      const startRow = parseInt(startAddress.match(/\d+/)?.[0] ?? "0", 10);
      const startColumn =
        (startAddress.match(/[A-Z]+/)?.[0] ?? "A").charCodeAt(0) - 65;
      const cellStyles: StylePresetsType = {};

      for (let i = 0; i < rows; i += 1) {
        for (let j = 0; j < columns; j += 1) {
          const cell = range.getCell(i, j);

          cell.load(["address", "numberFormat", "numberFormatLocal"]);
          cell.format.load([
            "fill/color",
            "font",
            "borders",
            "protection",
            "horizontalAlignment",
            "verticalAlignment",
            "wrapText",
            "indentLevel",
            "readingOrder",
            "textOrientation",
          ]);
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

          const borderEdges: EdgesType = [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
          ];
          const borders: BorderStyleType = {};

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
            alignment: {
              horizontalAlignment: cell.format.horizontalAlignment,
              verticalAlignment: cell.format.verticalAlignment,
              wrapText: cell.format.wrapText,
              indentLevel: cell.format.indentLevel,
              readingOrder: cell.format.readingOrder,
              textOrientation: cell.format.textOrientation,
            },
            numberFormat: cell.numberFormat[0][0],
            numberFormatLocal: cell.numberFormatLocal[0][0],
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
                weight: "Thin",
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

      allStylePresets[styleName] = cellStyles;

      await OfficeRuntime.storage.setItem(
        "cellStylePresets",
        JSON.stringify(allStylePresets),
      );

      const successMessage = {
        type: "success",
        title: "저장 완료",
        body: "선택한 셀 서식을 저장했습니다.",
      };

      updateState("setMessageList", successMessage);
    });
  } catch (error) {
    const errorMessage = {
      type: "error",
      title: "오류 발생",
      body: "셀 서식을 저장하는 중 오류가 발생했습니다.",
    };

    updateState("setMessageList", errorMessage);

    throw new Error(`Error in saveCellStylePreset: ${error}`);
  }
}

async function loadCellStylePreset(styleName: string) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "rowCount", "columnCount"]);
      await context.sync();

      const cellStylePresets: { [key: string]: StylePresetsType } = JSON.parse(
        await OfficeRuntime.storage.getItem("cellStylePresets"),
      );
      const savedCellStyles: StylePresetsType = cellStylePresets[styleName];

      if (!savedCellStyles) {
        throw new Error("Preset not found");
      }

      const rows = range.rowCount;
      const columns = range.columnCount;
      const savedAddresses = Object.keys(savedCellStyles);

      const savedRows = savedAddresses.reduce((max, addr) => {
        const row = parseInt(addr.match(/\d+$/)?.[0] ?? "0", 10);

        return Math.max(max, row);
      }, 0);

      const savedColumns = savedAddresses.reduce((max, addr) => {
        const col = addr.match(/^[A-Z]+/)?.[0] ?? "A";

        return Math.max(max, col.charCodeAt(0) - 64);
      }, 0);

      const getAddress = (row: number, column: number) => {
        const col = String.fromCharCode(64 + column);

        return `${col}${row}`;
      };

      const getStyle = (row: number, column: number) => {
        const address = getAddress(row, column);

        return savedCellStyles[address] || {};
      };

      const applyStyle = (cell: Excel.Range, styles: CellStyleType) => {
        if (styles.font) {
          const fontKeys: Array<keyof typeof styles.font> = [
            "name",
            "bold",
            "color",
            "size",
            "italic",
            "underline",
            "strikethrough",
            "tintAndShade",
          ];

          fontKeys.forEach((key) => {
            if (styles.font[key] !== undefined) {
              (cell.format.font[key] as Excel.RangeFont[typeof key]) =
                styles.font[key];
            }
          });
        }

        if (styles.fill && styles.fill.color) {
          cell.format.fill.color = styles.fill.color;
        }

        if (styles.borders) {
          const borders: ["EdgeBottom", "EdgeLeft", "EdgeTop", "EdgeRight"] = [
            "EdgeBottom",
            "EdgeLeft",
            "EdgeTop",
            "EdgeRight",
          ];

          borders.forEach((edge) => {
            if (styles.borders[edge]) {
              const border = cell.format.borders.getItem(edge);

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
          const alignmentProperties: Array<keyof typeof styles.alignment> = [
            "horizontalAlignment",
            "verticalAlignment",
            "wrapText",
            "indentLevel",
            "readingOrder",
            "textOrientation",
          ];

          alignmentProperties.forEach((prop) => {
            if (styles.alignment?.[prop] !== undefined) {
              (cell.format[prop] as Excel.RangeFormat[typeof prop]) =
                styles.alignment[prop];
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
          cell.numberFormat[0][0] = styles.numberFormat;
          cell.numberFormatLocal[0][0] = styles.numberFormatLocal;
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
    updateState("setMessageList", {
      type: "warning",
      title: "로드 오류",
      body: "프리셋을 불러오는데 실패하였습니다",
    });
  }
}

async function saveChartStylePreset(targetPreset: string, styleName: string) {
  try {
    if (styleName === "") {
      updateState("setMessageList", {
        type: "warning",
        title: "접근 오류",
        body: "프리셋을 정확하게 선택해주세요.",
      });

      return;
    }

    await Excel.run(async (context) => {
      let chartStylePresets: string | { [key: string]: ChartStyleType } =
        await OfficeRuntime.storage.getItem(targetPreset);

      if (typeof chartStylePresets === "string") {
        chartStylePresets = JSON.parse(chartStylePresets);
      }

      const selectedChart = context.workbook.getActiveChart();

      if (!selectedChart) {
        updateState("setMessageList", {
          type: "warning",
          title: "접근 오류",
          body: "선택된 차트를 찾을 수 없습니다.",
        });
      }

      selectedChart.load("chartType");
      await context.sync();

      const currentChartType = selectedChart.chartType;

      const propertiesToLoad = [
        "format/fill",
        "format/font",
        "format/border",
        "format/roundedCorners",
        "plotArea/format/fill",
        "plotArea/format/border",
        "plotArea/position",
        "plotArea/height",
        "plotArea/left",
        "plotArea/top",
        "plotArea/width",
        "plotArea/insideHeight",
        "plotArea/insideLeft",
        "plotArea/insideTop",
        "plotArea/insideWidth",
        "legend/format/fill",
        "legend/format/font",
        "legend/format/border",
        "legend/position",
        "seriesNameLevel",
      ];

      switch (currentChartType) {
        case Excel.ChartType.columnClustered:
        case Excel.ChartType.columnStacked:
        case Excel.ChartType.columnStacked100:
        case Excel.ChartType.line:
        case Excel.ChartType.lineStacked:
        case Excel.ChartType.lineStacked100:
        case Excel.ChartType.area:
        case Excel.ChartType.areaStacked:
        case Excel.ChartType.areaStacked100:
        case Excel.ChartType.histogram:
        case Excel.ChartType.boxwhisker:
        case Excel.ChartType.waterfall:
        case Excel.ChartType.funnel:
        case Excel.ChartType._3DArea:
        case Excel.ChartType._3DAreaStacked:
        case Excel.ChartType._3DAreaStacked100:
        case Excel.ChartType._3DColumn:
        case Excel.ChartType._3DColumnClustered:
        case Excel.ChartType._3DColumnStacked:
        case Excel.ChartType._3DColumnStacked100:
        case Excel.ChartType._3DLine:
        case Excel.ChartType._3DBarClustered:
        case Excel.ChartType._3DBarStacked:
        case Excel.ChartType._3DBarStacked100:
          propertiesToLoad.push(
            "axes/categoryAxis/format/line",
            "axes/categoryAxis/format/font",
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/categoryAxis/position",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.pie:
        case Excel.ChartType.doughnut:
        case Excel.ChartType.treemap:
        case Excel.ChartType.sunburst:
        case Excel.ChartType._3DPie:
        case Excel.ChartType._3DPieExploded:
          propertiesToLoad.push("series");
          break;

        case Excel.ChartType.bubble:
        case Excel.ChartType.xyscatter:
        case Excel.ChartType.xyscatterLines:
        case Excel.ChartType.xyscatterLinesNoMarkers:
        case Excel.ChartType.xyscatterSmooth:
        case Excel.ChartType.xyscatterSmoothNoMarkers:
          propertiesToLoad.push(
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.stockHLC:
        case Excel.ChartType.stockOHLC:
        case Excel.ChartType.stockVHLC:
        case Excel.ChartType.stockVOHLC:
        case Excel.ChartType.surface:
        case Excel.ChartType.surfaceTopView:
        case Excel.ChartType.surfaceTopViewWireframe:
        case Excel.ChartType.surfaceWireframe:
          propertiesToLoad.push(
            "axes/categoryAxis/format/line",
            "axes/categoryAxis/format/font",
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/categoryAxis/position",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.radar:
        case Excel.ChartType.radarFilled:
        case Excel.ChartType.radarMarkers:
          propertiesToLoad.push(
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.regionMap:
          propertiesToLoad.push("series");
          break;

        case Excel.ChartType.barClustered:
        case Excel.ChartType.barStacked:
        case Excel.ChartType.barStacked100:
        case Excel.ChartType.coneBarClustered:
        case Excel.ChartType.coneBarStacked:
        case Excel.ChartType.coneBarStacked100:
        case Excel.ChartType.cylinderBarClustered:
        case Excel.ChartType.cylinderBarStacked:
        case Excel.ChartType.cylinderBarStacked100:
        case Excel.ChartType.pyramidBarClustered:
        case Excel.ChartType.pyramidBarStacked:
        case Excel.ChartType.pyramidBarStacked100:
        case Excel.ChartType.barOfPie:
          propertiesToLoad.push(
            "axes/categoryAxis/format/line",
            "axes/categoryAxis/format/font",
            "axes/valueAxis/format/line",
            "axes/valueAxis/format/font",
            "series",
            "axes/categoryAxis/position",
            "axes/valueAxis/position",
          );
          break;

        case Excel.ChartType.lineMarkers:
        case Excel.ChartType.lineMarkersStacked:
        case Excel.ChartType.lineMarkersStacked100:
        case Excel.ChartType.pareto:
        case Excel.ChartType.pieExploded:
        case Excel.ChartType.pieOfPie:
        case Excel.ChartType.doughnutExploded:
          propertiesToLoad.push("series");
          break;

        default:
          updateState("setMessageList", {
            type: "warning",
            title: "접근 오류",
            body: "지원하지 않는 차트 유형입니다.",
          });

          return;
      }

      selectedChart.load(propertiesToLoad);
      await context.sync();

      const chartFillColor = selectedChart.format.fill.getSolidColor();
      const legendFillColor = selectedChart.legend.format.fill.getSolidColor();
      const plotAreaFillColor =
        selectedChart.plotArea.format.fill.getSolidColor();

      await context.sync();

      const chartStyle: ChartStyleType = {
        chartType: currentChartType,
        font: {
          name: selectedChart.format.font.name,
          size: selectedChart.format.font.size,
          color: selectedChart.format.font.color,
          bold: selectedChart.format.font.bold,
          italic: selectedChart.format.font.italic,
          underline: selectedChart.format.font.underline,
        },
        roundedCorners: selectedChart.format.roundedCorners,
        fill: {
          color: chartFillColor.value,
        },
        border: {
          lineStyle: selectedChart.format.border.lineStyle,
          color: selectedChart.format.border.color,
          weight: selectedChart.format.border.weight,
        },
        plotArea: {
          fill: plotAreaFillColor.value,
          border: {
            lineStyle: selectedChart.plotArea.format.border.lineStyle,
            color: selectedChart.plotArea.format.border.color,
            weight: selectedChart.plotArea.format.border.weight,
          },
          position: selectedChart.plotArea.position,
          height: selectedChart.plotArea.height,
          left: selectedChart.plotArea.left,
          top: selectedChart.plotArea.top,
          width: selectedChart.plotArea.width,
          insideHeight: selectedChart.plotArea.insideHeight,
          insideLeft: selectedChart.plotArea.insideLeft,
          insideTop: selectedChart.plotArea.insideTop,
          insideWidth: selectedChart.plotArea.insideWidth,
        },
        legend: {
          fill: legendFillColor.value,
          font: {
            name: selectedChart.legend.format.font.name,
            size: selectedChart.legend.format.font.size,
            color: selectedChart.legend.format.font.color,
            bold: selectedChart.legend.format.font.bold,
            italic: selectedChart.legend.format.font.italic,
            underline: selectedChart.legend.format.font.underline,
          },
          border: {
            lineStyle: selectedChart.legend.format.border.lineStyle,
            color: selectedChart.legend.format.border.color,
            weight: selectedChart.legend.format.border.weight,
          },
          position: selectedChart.legend.position,
        },
        seriesNameLevel: selectedChart.seriesNameLevel,
        axes: {},
        series: [],
      };

      if (propertiesToLoad.includes("axes/categoryAxis")) {
        chartStyle.axes.categoryAxis = {
          position: selectedChart.axes.categoryAxis.position,
          format: {
            line: {
              color: selectedChart.axes.categoryAxis.format.line.color,
              style: selectedChart.axes.categoryAxis.format.line.lineStyle,
              weight: selectedChart.axes.categoryAxis.format.line.weight,
            },
            font: {
              name: selectedChart.axes.categoryAxis.format.font.name,
              size: selectedChart.axes.categoryAxis.format.font.size,
              color: selectedChart.axes.categoryAxis.format.font.color,
              bold: selectedChart.axes.categoryAxis.format.font.bold,
              italic: selectedChart.axes.categoryAxis.format.font.italic,
              underline: selectedChart.axes.categoryAxis.format.font.underline,
            },
          },
        };
      }

      if (propertiesToLoad.includes("axes/valueAxis")) {
        chartStyle.axes.valueAxis = {
          position: selectedChart.axes.valueAxis.position,
          format: {
            line: {
              color: selectedChart.axes.valueAxis.format.line.color,
              style: selectedChart.axes.valueAxis.format.line.lineStyle,
              weight: selectedChart.axes.valueAxis.format.line.weight,
            },
            font: {
              name: selectedChart.axes.valueAxis.format.font.name,
              size: selectedChart.axes.valueAxis.format.font.size,
              color: selectedChart.axes.valueAxis.format.font.color,
              bold: selectedChart.axes.valueAxis.format.font.bold,
              italic: selectedChart.axes.valueAxis.format.font.italic,
              underline: selectedChart.axes.valueAxis.format.font.underline,
            },
          },
        };
      }

      if (propertiesToLoad.includes("series")) {
        selectedChart.series.load("items");
        await context.sync();

        chartStyle.series = [];

        for (let i = 0; i < selectedChart.series.items.length; i += 1) {
          const series = selectedChart.series.items[i];

          series.load(["format/fill", "format/line"]);
          await context.sync();

          chartStyle.series.push(series);
        }
      }

      if (typeof chartStylePresets !== "string") {
        chartStylePresets[styleName] = chartStyle;
      }

      await OfficeRuntime.storage.setItem(
        targetPreset,
        JSON.stringify(chartStylePresets),
      );

      updateState("setMessageList", {
        type: "success",
        title: "저장 성공",
        body: "차트 서식 프리셋이 저장되었습니다.",
      });
    });
  } catch (error) {
    if (error instanceof Error) {
      updateState("setMessageList", {
        type: "error",
        title: "오류 발생",
        body: `차트 서식 프리셋을 저장하는 중 오류가 발생했습니다. ${error.message}`,
      });
      throw new Error(error.message);
    }
  }
}

async function loadChartStylePreset(targetPreset: string, styleName: string) {
  if (styleName === "") {
    updateState("setMessageList", {
      type: "warning",
      title: "접근 오류",
      body: "프리셋을 정확하게 선택해주세요",
    });

    return;
  }

  try {
    await Excel.run(async (context) => {
      const currentChart = context.workbook.getActiveChart();

      if (!currentChart) {
        updateState("setMessageList", {
          type: "error",
          title: "차트 없음",
          body: "선택된 차트를 찾을 수 없습니다.",
        });

        return;
      }

      currentChart.load("chartType");
      await context.sync();

      let chartStylePresets: { [key: string]: ChartStyleType } =
        await JSON.parse(OfficeRuntime.storage.getItem(targetPreset));

      if (!chartStylePresets) {
        updateState("setMessageList", {
          type: "warning",
          title: "로드 실패",
          body: "프리셋 목록을 불러오는데 실패했습니다.",
        });

        return;
      }

      const chartStyle = chartStylePresets[styleName];

      if (!chartStyle) {
        updateState("setMessageList", {
          type: "warning",
          title: "접근 오류",
          body: "해당 프리셋을 찾을 수 없습니다.",
        });

        return;
      }

      if (
        chartStyle.chartType &&
        currentChart.chartType !== chartStyle.chartType
      ) {
        updateState("setMessageList", {
          type: "warning",
          title: "차트 유형 불일치",
          body: "저장된 스타일의 차트 유형과 현재 차트의 유형이 다릅니다. 일부 스타일이 적용되지 않을 수 있습니다.",
        });
      }

      applyBasicChartProperties(currentChart, chartStyle);
      applyLegendProperties(currentChart, chartStyle);
      applyPlotAreaProperties(currentChart, chartStyle);
      applyAxisProperties(currentChart, chartStyle);

      if (chartStyle.series && currentChart.series) {
        context.trackedObjects.add(currentChart);

        await applySeriesProperties(currentChart, chartStyle);

        context.trackedObjects.remove(currentChart);
      }

      await context.sync();

      updateState("setMessageList", {
        type: "success",
        title: "적용 성공",
        body: "차트 서식을 적용했습니다.",
      });
    });
  } catch (error) {
    updateState("setMessageList", {
      type: "error",
      title: "적용 실패",
      body: `차트 서식 적용에 실패하였습니다.`,
    });
  }
}

function applyBasicChartProperties(
  currentChart: Excel.Chart,
  chartStyle: ChartStyleType,
) {
  if (chartStyle.fill.color) {
    currentChart.format.fill.setSolidColor(chartStyle.fill.color.m_value);
  } else {
    currentChart.format.fill.clear();
  }

  if (chartStyle.border) {
    if (chartStyle.border.lineStyle !== "none") {
      if (chartStyle.border.color) {
        currentChart.format.border.color = chartStyle.border.color;
      }

      if (chartStyle.border.lineStyle) {
        currentChart.format.border.lineStyle = chartStyle.border.lineStyle;
      }

      if (chartStyle.border.weight && chartStyle.border.weight > 0) {
        currentChart.format.border.weight = chartStyle.border.weight;
      }
    } else {
      currentChart.format.border.clear();
    }
  }

  if (chartStyle.font) {
    const chartFontKeys: Array<keyof typeof chartStyle.font> = [
      "name",
      "size",
      "color",
      "bold",
      "italic",
      "underline",
    ];

    chartFontKeys.forEach((key) => {
      if (chartStyle.font[key] !== undefined) {
        (currentChart.format.font[key] as Excel.ChartFont[typeof key]) =
          chartStyle.font[key];
      }
    });
  }

  if (chartStyle.roundedCorners !== undefined) {
    currentChart.format.roundedCorners = chartStyle.roundedCorners;
  }
}

function applyLegendProperties(
  currentChart: Excel.Chart,
  chartStyle: ChartStyleType,
) {
  if (chartStyle.legend) {
    if (chartStyle.legend.fill) {
      currentChart.legend.format.fill.setSolidColor(chartStyle.legend.fill);
    } else {
      currentChart.legend.format.fill.clear();
    }

    if (chartStyle.legend.border) {
      if (chartStyle.border.lineStyle !== "None") {
        if (chartStyle.legend.border.color) {
          currentChart.legend.format.border.color =
            chartStyle.legend.border.color;
        }

        if (chartStyle.legend.border.lineStyle) {
          currentChart.legend.format.border.lineStyle =
            chartStyle.legend.border.lineStyle;
        }

        if (chartStyle.border.weight && chartStyle.border.weight > 0) {
          currentChart.legend.format.border.weight =
            chartStyle.legend.border.weight;
        }
      } else {
        currentChart.legend.format.border.clear();
      }
    }

    if (chartStyle.legend.font) {
      const legendFontKeys: Array<keyof typeof chartStyle.legend.font> = [
        "name",
        "size",
        "color",
        "bold",
        "italic",
        "underline",
      ];

      legendFontKeys.forEach((key) => {
        if (chartStyle.legend.font[key] !== undefined) {
          (currentChart.legend.format.font[
            key
          ] as Excel.ChartFont[typeof key]) = chartStyle.legend.font[key];
        }
      });
    }

    if (chartStyle.legend.position) {
      currentChart.legend.position = chartStyle.legend.position;
    }
  }
}

function applyPlotAreaProperties(
  currentChart: Excel.Chart,
  chartStyle: ChartStyleType,
) {
  if (chartStyle.plotArea) {
    if (chartStyle.plotArea.fill) {
      currentChart.plotArea.format.fill.setSolidColor(chartStyle.plotArea.fill);
    } else {
      currentChart.plotArea.format.fill.clear();
    }

    if (chartStyle.plotArea.border) {
      if (chartStyle.plotArea.border.lineStyle !== "None") {
        if (chartStyle.plotArea.border.color) {
          currentChart.plotArea.format.border.color =
            chartStyle.plotArea.border.color;
        }

        if (chartStyle.plotArea.border.lineStyle) {
          currentChart.plotArea.format.border.lineStyle =
            chartStyle.plotArea.border.lineStyle;
        }

        if (
          chartStyle.plotArea.border.weight &&
          chartStyle.plotArea.border.weight > 0
        ) {
          currentChart.plotArea.format.border.weight =
            chartStyle.plotArea.border.weight;
        }
      } else {
        currentChart.plotArea.format.border.clear();
      }
    }

    if (chartStyle.plotArea.position === "Automatic") {
      currentChart.plotArea.position = chartStyle.plotArea.position;
    } else {
      currentChart.plotArea.height = chartStyle.plotArea.height;
      currentChart.plotArea.left = chartStyle.plotArea.left;
      currentChart.plotArea.top = chartStyle.plotArea.top;
      currentChart.plotArea.width = chartStyle.plotArea.width;
      currentChart.plotArea.insideHeight = chartStyle.plotArea.insideHeight;
      currentChart.plotArea.insideLeft = chartStyle.plotArea.insideLeft;
      currentChart.plotArea.insideTop = chartStyle.plotArea.insideTop;
      currentChart.plotArea.insideWidth = chartStyle.plotArea.insideWidth;
    }
  }
}

function applyAxisProperties(
  currentChart: Excel.Chart,
  chartStyle: ChartStyleType,
) {
  if (chartStyle.axes) {
    if (chartStyle.axes.categoryAxis && currentChart.axes.categoryAxis) {
      applySingleAxisProperties(
        currentChart.axes.categoryAxis,
        chartStyle.axes.categoryAxis,
      );
    }

    if (chartStyle.axes.valueAxis && currentChart.axes.valueAxis) {
      applySingleAxisProperties(
        currentChart.axes.valueAxis,
        chartStyle.axes.valueAxis,
      );
    }
  }
}

function applySingleAxisProperties(
  axis: Excel.ChartAxis,
  axisStyle: ValueAxisType,
) {
  if (axisStyle.format) {
    if (axisStyle.format.line.style !== "None") {
      if (axisStyle.format.line) {
        axis.format.line.color = axisStyle.format.line.color;
        axis.format.line.lineStyle = axisStyle.format.line.style;

        if (axisStyle.format.line.weight > 0) {
          axis.format.line.weight = axisStyle.format.line.weight;
        }
      }
    } else {
      axis.format.line.clear();
    }

    if (axisStyle.format.font) {
      const axisFontKeys: Array<keyof typeof axisStyle.format.font> = [
        "name",
        "size",
        "color",
        "bold",
        "italic",
        "underline",
      ];

      axisFontKeys.forEach((key) => {
        if (axisStyle.format.font[key] !== undefined)
          (axis.format.font[key] as Excel.ChartFont[typeof key]) =
            axisStyle.format.font[key];
      });
    }
  }

  if (axisStyle.position) {
    axis.position = axisStyle.position;
  }
}

async function applySeriesProperties(
  currentChart: Excel.Chart,
  chartStyle: ChartStyleType,
) {
  if (chartStyle.series && currentChart.series) {
    await currentChart.series.load("items");
    await currentChart.context.sync();

    const seriesArray: Excel.ChartSeries[] = Array.isArray(chartStyle.series)
      ? chartStyle.series
      : Object.values(chartStyle.series);

    for (
      let index = 0;
      index < Math.min(seriesArray.length, currentChart.series.items.length);
      index += 1
    ) {
      const series = currentChart.series.items[index];

      if (series) {
        series.load(["format/fill", "format/line"]);
      }
    }

    await currentChart.context.sync();

    for (
      let index = 0;
      index < Math.min(seriesArray.length, currentChart.series.items.length);
      index += 1
    ) {
      const seriesStyle = seriesArray[index];
      const series = currentChart.series.items[index];

      if (seriesStyle.format) {
        if (seriesStyle.format.fill) {
          series.format.fill.setSolidColor(
            seriesStyle.format.fill.getSolidColor().value,
          );
        } else {
          series.format.fill.clear();
        }
      }

      if (seriesStyle.format.line.lineStyle !== "None") {
        if (seriesStyle.format.line.color) {
          series.format.line.color = seriesStyle.format.line.color;
        }

        if (seriesStyle.format.line.lineStyle) {
          series.format.line.lineStyle = seriesStyle.format.line.lineStyle;
        }

        if (
          seriesStyle.format.line.weight &&
          seriesStyle.format.line.weight > 0
        ) {
          series.format.line.weight = seriesStyle.format.line.weight;
        }
      } else {
        series.format.line.clear();
      }
    }
  }
}

export {
  storeCellStyle,
  applyCellStyle,
  highlightingCell,
  changeCellborder,
  saveCellStylePreset,
  loadCellStylePreset,
  saveChartStylePreset,
  loadChartStylePreset,
  detectErrorCell,
};
