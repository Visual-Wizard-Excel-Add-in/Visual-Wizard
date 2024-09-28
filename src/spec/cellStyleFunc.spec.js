import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  storeCellStyle,
  applyCellStyle,
  detectErrorCell,
  highlightingCell,
  saveCellStylePreset,
  loadCellStylePreset,
} from "../taskpane/utils/cellStyleFuncs";
import {
  saveChartStylePreset,
  loadChartStylePreset,
} from "../taskpane/utils/chartStyleFuncs";
import * as cellCommonUtils from "../taskpane/utils/commonFuncs";

const mockChart = {
  load: vi.fn(),
  context: { sync: vi.fn() },
  chartType: "ColumnClustered",
  format: {
    fill: {
      getSolidColor: vi.fn().mockReturnValue({ m_value: "red" }),
      setSolidColor: vi.fn(),
    },
    font: {
      name: "Arial",
      size: 12,
      color: "black",
      bold: true,
      italic: false,
      underline: "Single",
    },
    border: {
      color: "black",
      weight: 1,
      lineStyle: "Continuous",
    },
    roundedCorners: false,
  },
  legend: {
    format: {
      fill: {
        getSolidColor: vi.fn().mockReturnValue({ m_value: "blue" }),
        setSolidColor: vi.fn(),
      },
      font: {
        name: "Arial",
        size: 12,
        color: "black",
        bold: false,
        italic: false,
        underline: "None",
      },
      border: {
        color: "black",
        weight: 1,
        lineStyle: "Continuous",
      },
    },
    position: "right",
  },
  plotArea: {
    format: {
      fill: {
        getSolidColor: vi.fn().mockReturnValue({ m_value: "white" }),
        setSolidColor: vi.fn(),
      },
      font: {
        name: "Arial",
        size: 12,
        color: "black",
        bold: true,
        italic: false,
        underline: "Single",
      },
      border: {
        color: "black",
        weight: 1,
        lineStyle: "Continuous",
      },
    },
  },
  axes: {
    categoryAxis: {
      format: {
        font: {
          name: "Arial",
        },
        line: {
          color: "black",
          weight: 1,
          lineStyle: "Continuous",
        },
      },
      position: "bottom",
    },
    valueAxis: {
      format: {
        font: {
          name: "Arial",
        },
        line: {
          color: "black",
          weight: 1,
          lineStyle: "Continuous",
        },
      },
      position: "left",
    },
  },
  series: {
    load: vi.fn(),
    items: [
      {
        load: vi.fn(),
        format: {
          fill: {
            getSolidColor: vi.fn().mockReturnValue({ m_value: "blue" }),
            setSolidColor: vi.fn(),
          },
          font: {
            name: "Arial",
            size: 12,
            color: "black",
            bold: true,
            italic: false,
            underline: "Single",
          },
          line: {
            color: "black",
            weight: 1,
            lineStyle: "Continuous",
          },
        },
      },
    ],
  },
};

global.Excel = {
  run: vi.fn(async (callback) => {
    const context = {
      workbook: {
        worksheets: {
          getActiveWorksheet: vi.fn().mockReturnValue({
            getRange: vi.fn().mockReturnValue({
              load: vi.fn(),
              context: {
                sync: vi.fn(),
              },
              format: {
                borders: {
                  getItem: vi.fn().mockReturnValue({
                    load: vi.fn(),
                    color: "black",
                    style: "Continuous",
                    weight: 2,
                  }),
                },
                fill: {
                  color: "red",
                },
                font: {
                  name: "Arial",
                  bold: true,
                  color: "black",
                  italic: false,
                  size: 12,
                  underline: "Single",
                  tintAndShade: 0.5,
                },
                horizontalAlignment: "center",
                verticalAlignment: "middle",
              },
              numberFormat: "General",
            }),
            getUsedRange: vi.fn().mockReturnValue({
              load: vi.fn(),
              values: [["#DIV/0!", "#N/A"]],
              getCell: vi.fn().mockReturnValue({
                load: vi.fn(),
                format: {
                  fill: { color: "" },
                  borders: {
                    getItem: vi.fn().mockReturnValue({
                      color: "",
                      style: "",
                      weight: 0,
                    }),
                  },
                },
              }),
            }),
          }),
        },
        getActiveChart: vi.fn().mockReturnValue(mockChart),
      },
      sync: vi.fn(),
      trackedObjects: {
        add: vi.fn(),
        remove: vi.fn(),
      },
    };

    await callback(context);

    return context.workbook;
  }),
  ChartType: {
    columnClustered: "ColumnClustered",
  },
  BorderLineStyle: {
    none: "None",
  },
  BorderWeight: {
    thick: "Thick",
  },
};

global.OfficeRuntime = {
  storage: {
    getItem: vi.fn(),
    setItem: vi.fn(),
  },
};

vi.mock("../taskpane/utils/cellCommonUtils", () => ({
  updateState: vi.fn(),
  extractArgsAddress: vi.fn().mockImplementation((address) => address),
}));

describe("cellStyleFunc", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("saveChartStylePreset", () => {
    it("should save chart style preset", async () => {
      const styleName = "TestChartStyle";

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));
      OfficeRuntime.storage.setItem.mockResolvedValueOnce();

      await saveChartStylePreset("allChartPresets", styleName);

      expect(OfficeRuntime.storage.setItem).toHaveBeenCalledWith(
        "allChartPresets",
        expect.stringContaining(styleName),
      );
    });

    it("should display warning for empty style name", async () => {
      await saveChartStylePreset("allChartPresets", "");

      expect(cellCommonUtils.updateState).toHaveBeenCalledWith(
        "setMessageList",
        expect.objectContaining({
          type: "warning",
          title: expect.any(String),
          body: expect.stringContaining("프리셋을 정확하게 선택해주세요"),
        }),
      );
    });
  });

  describe("loadChartStylePreset", () => {
    it("should load and apply chart style preset", async () => {
      const styleName = "TestChartStyle";
      const mockPreset = {
        [styleName]: {
          chartType: "ColumnClustered",
          fill: { color: { m_value: "red" } },
          legend: {
            fill: { color: { m_value: "blue" } },
            position: "right",
          },
          plotArea: {
            fill: { color: { m_value: "white" } },
          },
          series: [
            {
              format: {
                fill: { color: "blue" },
              },
            },
          ],
        },
      };

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(
        JSON.stringify(mockPreset),
      );

      await loadChartStylePreset("allChartPresets", styleName);

      expect(mockChart.format.fill.setSolidColor).toHaveBeenCalledWith("red");
      expect(mockChart.legend.format.fill.setSolidColor).toHaveBeenCalledWith(
        "blue",
      );
      expect(mockChart.plotArea.format.fill.setSolidColor).toHaveBeenCalledWith(
        "white",
      );
      expect(
        mockChart.series.items[0].format.fill.setSolidColor,
      ).toHaveBeenCalledWith("blue");
    });

    it("should handle non-existent preset", async () => {
      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));

      await loadChartStylePreset("allChartPresets", "NonExistentStyle");

      expect(cellCommonUtils.updateState).toHaveBeenCalledWith(
        "setMessageList",
        expect.objectContaining({
          type: "warning",
          title: expect.any(String),
          body: expect.stringContaining("해당 프리셋을 찾을 수 없습니다"),
        }),
      );
    });
  });

  // 새로운 테스트 코드 추가

  describe("storeCellStyle", () => {
    it("should store cell style if isCellHighlighting is true", async () => {
      const cellAddress = "A1";
      const allPresets = "allCellStyles";

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));

      await storeCellStyle(cellAddress, allPresets, true);

      expect(OfficeRuntime.storage.setItem).toHaveBeenCalledWith(
        allPresets,
        expect.any(String),
      );
    });

    it("should return cell style if allPresets is allMacroPresets", async () => {
      const cellAddress = "A1";
      const allPresets = "allMacroPresets";

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));

      const result = await storeCellStyle(cellAddress, allPresets, true);

      expect(result).toHaveProperty("color");
      expect(result).toHaveProperty("borders");
    });
  });

  describe("applyCellStyle", () => {
    it("should apply the saved cell style to the cell", async () => {
      const cellAddress = "A1";
      const mockPreset = {
        [cellAddress]: {
          font: {
            bold: true,
            color: "black",
            italic: false,
            size: 12,
            underline: "Single",
          },
          color: "red",
          borders: {
            EdgeBottom: {
              color: "black",
              style: "Continuous",
              weight: 2,
            },
          },
          numberFormat: "General",
          horizontalAlignment: "center",
          verticalAlignment: "middle",
        },
      };

      OfficeRuntime.storage.getItem.mockResolvedValue(
        JSON.stringify(mockPreset),
      );

      await applyCellStyle(cellAddress, "cellStylePresets", false);

      expect(global.Excel.run).toHaveBeenCalled();

      const runCallback = global.Excel.run.mock.calls[0][0];

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: vi.fn().mockReturnValue({
              getRange: vi.fn().mockReturnValue({
                format: {
                  fill: { color: "" },
                  font: {},
                  borders: {
                    getItem: vi.fn().mockReturnValue({}),
                  },
                  horizontalAlignment: "",
                  verticalAlignment: "",
                },
                numberFormat: "",
              }),
            }),
          },
        },
        sync: vi.fn(),
      };

      await runCallback(mockContext);

      const cell = mockContext.workbook.worksheets
        .getActiveWorksheet()
        .getRange(cellAddress);

      expect(cell.format.fill.color).toBe("red");
      expect(cell.format.font.bold).toBe(true);
      expect(cell.format.font.color).toBe("black");
      expect(cell.format.font.size).toBe(12);
      expect(cell.format.font.underline).toBe("Single");
      expect(cell.format.horizontalAlignment).toBe("center");
      expect(cell.format.verticalAlignment).toBe("middle");
      expect(cell.format.borders.getItem("EdgeBottom").color).toBe("black");
      expect(cell.format.borders.getItem("EdgeBottom").style).toBe(
        "Continuous",
      );
      expect(cell.format.borders.getItem("EdgeBottom").weight).toBe(2);
    });

    it("should throw an error if no saved presets are found", async () => {
      OfficeRuntime.storage.getItem.mockResolvedValueOnce(null);

      await expect(
        applyCellStyle("A1", "cellStylePresets", false),
      ).rejects.toThrow("No any saved presets found.");
    });

    it("should remove the applied style from storage after applying", async () => {
      const cellAddress = "A1";
      const mockPreset = {
        [cellAddress]: {
          font: { bold: true, color: "black" },
          color: "red",
        },
      };

      global.OfficeRuntime.storage.getItem.mockResolvedValue(
        JSON.stringify(mockPreset),
      );

      await applyCellStyle(cellAddress, "cellStylePresets", false);

      expect(global.OfficeRuntime.storage.setItem).toHaveBeenCalledWith(
        "cellStylePresets",
        JSON.stringify({}),
      );
    });

    it("should handle applying a style with missing properties", async () => {
      const cellAddress = "A1";
      const mockPreset = {
        [cellAddress]: {
          font: { bold: true },
        },
      };

      OfficeRuntime.storage.getItem.mockResolvedValue(
        JSON.stringify(mockPreset),
      );

      await applyCellStyle(cellAddress, "cellStylePresets", false);

      const workbook = await global.Excel.run.mock.results[0].value;
      const sheet = workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);

      expect(cell.format.font.bold).toBe(true);
      expect(cell.format.fill.color).toBe("red");
    });

    it("should throw an error if no presets found", async () => {
      const cellAddress = "A1";
      const allPresets = "allCellStyles";

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(null);

      await expect(
        applyCellStyle(cellAddress, allPresets, false),
      ).rejects.toThrow("No any saved presets found.");
    });
  });

  describe("detectErrorCell", () => {
    it("should highlight error cells", async () => {
      const isCellHighlighting = true;

      await detectErrorCell(isCellHighlighting);

      expect(global.Excel.run).toHaveBeenCalled();

      const runCallback = global.Excel.run.mock.calls[0][0];

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: vi.fn().mockReturnValue({
              getRange: vi.fn().mockReturnValue({
                format: {
                  fill: { color: "" },
                  font: {},
                  borders: {
                    getItem: vi.fn().mockReturnValue({}),
                  },
                  horizontalAlignment: "",
                  verticalAlignment: "",
                },
                numberFormat: "",
              }),
              getUsedRange: vi.fn().mockReturnValue({
                load: vi.fn(),
                values: [["#DIV/0!", "#N/A"]],
                getCell: vi.fn().mockReturnValue({
                  load: vi.fn(),
                  format: {
                    fill: { color: "" },
                    borders: {
                      getItem: vi.fn().mockReturnValue({
                        color: "",
                        style: "",
                        weight: 0,
                      }),
                    },
                  },
                }),
              }),
            }),
          },
        },
        sync: vi.fn(),
      };

      await runCallback(mockContext);

      const range = mockContext.workbook.worksheets
        .getActiveWorksheet()
        .getUsedRange();

      const errorCells = range.getCell.mock.calls.filter(
        ([rowIndex, colIndex]) =>
          [
            "#DIV/0!",
            "#N/A",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#NULL!",
          ].includes(range.values[rowIndex][colIndex]),
      );

      errorCells.forEach(([rowIndex, colIndex]) => {
        const cell = range.getCell(rowIndex, colIndex);
        expect(cell.format.fill.color).toBe("red");
        expect(cell.format.borders.getItem).toHaveBeenCalledWith("EdgeBottom");
      });
    });
  });

  describe("highlightingCell", () => {
    it("should highlight cells if isCellHighlighting is true", async () => {
      const isCellHighlighting = true;
      const argCells = ["A1", "B1"];
      const resultCell = "C1";

      await highlightingCell(isCellHighlighting, argCells, resultCell);

      const runCallback = global.Excel.run.mock.calls[0][0];

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: vi.fn().mockReturnValue({
              getRange: vi.fn().mockReturnValue({
                context: { sync: vi.fn() },
                format: {
                  fill: { color: "" },
                  font: {},
                  borders: {
                    getItem: vi.fn().mockReturnValue({}),
                  },
                  horizontalAlignment: "",
                  verticalAlignment: "",
                },
                numberFormat: "",
              }),
            }),
          },
        },
        sync: vi.fn(),
      };

      await runCallback(mockContext);

      const firstArgcellRange = mockContext.workbook.worksheets
        .getActiveWorksheet()
        .getRange(argCells[0]);
      const secondArgcellRange = mockContext.workbook.worksheets
        .getActiveWorksheet()
        .getRange(argCells[1]);

      expect(firstArgcellRange.format.fill.color).toBe("#28f925");
      expect(secondArgcellRange.format.fill.color).toBe("#28f925");
    });

    it("should apply cell style if isCellHighlighting is false", async () => {
      const isCellHighlighting = false;
      const argCells = ["A1", "B1"];
      const resultCell = "C1";

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));

      await highlightingCell(isCellHighlighting, argCells, resultCell);

      const runCallback = global.Excel.run.mock.calls[0][0];

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: vi.fn().mockReturnValue({
              getRange: vi.fn().mockReturnValue({
                context: { sync: vi.fn() },
                format: {
                  fill: { color: "" },
                  font: {},
                  borders: {
                    getItem: vi.fn().mockReturnValue({}),
                  },
                  horizontalAlignment: "",
                  verticalAlignment: "",
                },
                numberFormat: "",
              }),
            }),
          },
        },
        sync: vi.fn(),
      };

      await runCallback(mockContext);

      const resultCellRange = mockContext.workbook.worksheets
        .getActiveWorksheet()
        .getRange(resultCell);

      expect(resultCellRange.format.fill.color).toBe("");
    });
  });

  describe("saveCellStylePreset", () => {
    it("should save the cell style preset", async () => {
      const styleName = "TestStyle";

      global.Excel.run.mockImplementation(async (callback) => {
        const context = {
          workbook: {
            getSelectedRange: () => ({
              load: vi.fn(),
              rowCount: 1,
              columnCount: 1,
              address: "Sheet1!A1",
              getCell: () => ({
                address: "Sheet1!A1",
                load: vi.fn(),
                format: {
                  load: vi.fn(),
                  font: {
                    load: vi.fn(),
                    name: "Arial",
                    size: 11,
                    color: "#000000",
                    bold: false,
                    italic: false,
                    underline: "None",
                    strikethrough: false,
                  },
                  fill: { color: "#FFFFFF" },
                  borders: {
                    getItem: () => ({
                      load: vi.fn(),
                      style: "Continuous",
                      color: "#000000",
                      weight: "Thin",
                    }),
                  },
                  protection: {
                    load: vi.fn(),
                    locked: false,
                    formulaHidden: false,
                  },
                  horizontalAlignment: "General",
                  verticalAlignment: "Bottom",
                  wrapText: false,
                  indentLevel: 0,
                  readingOrder: "LeftToRight",
                  textOrientation: 0,
                },
                numberFormat: "General",
                numberFormatLocal: "General",
              }),
            }),
          },
          sync: vi.fn(),
        };
        await callback(context);
      });

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));
      global.OfficeRuntime.storage.setItem.mockResolvedValue(undefined);

      global.updateState = vi.fn();

      await saveCellStylePreset(styleName);

      expect(OfficeRuntime.storage.setItem).toHaveBeenCalledWith(
        "cellStylePresets",
        expect.any(String),
      );
      expect(cellCommonUtils.updateState).toHaveBeenCalledWith(
        "setMessageList",
        expect.objectContaining({
          type: "success",
          title: "저장 완료",
          body: "선택한 셀 서식을 저장했습니다.",
        }),
      );
    });

    it("should display a warning if the style name is empty", async () => {
      await saveCellStylePreset("");

      expect(cellCommonUtils.updateState).toHaveBeenCalledWith(
        "setMessageList",
        expect.objectContaining({
          type: "warning",
          title: "접근 오류",
          body: "프리셋을 정확하게 선택해주세요",
        }),
      );
    });
  });

  describe("loadCellStylePreset", () => {
    beforeEach(() => {
      vi.resetAllMocks();

      global.Excel.run = vi.fn(async (callback) => {
        const context = {
          workbook: {
            getSelectedRange: () => ({
              load: vi.fn(),
              address: "Sheet1!A1",
              rowCount: 1,
              columnCount: 1,
              getCell: () => ({
                format: {
                  font: { bold: false, color: "" },
                  fill: { color: "" },
                  borders: {
                    getItem: () => ({
                      style: "",
                      color: "",
                      weight: "",
                    }),
                  },
                  protection: {
                    locked: false,
                    formulaHidden: false,
                  },
                  horizontalAlignment: "",
                  verticalAlignment: "",
                  wrapText: false,
                  indentLevel: 0,
                  readingOrder: "",
                  textOrientation: 0,
                },
                numberFormat: "",
                numberFormatLocal: "",
              }),
            }),
          },
          sync: vi.fn(),
        };
        await callback(context);
      });

      global.OfficeRuntime = {
        storage: {
          getItem: vi.fn(),
          setItem: vi.fn(),
        },
      };

      global.updateState = vi.fn();
    });

    it("should load and apply cell style preset", async () => {
      const styleName = "TestStyle";
      const mockPreset = {
        A1: {
          font: { bold: true, color: "black" },
          fill: { color: "red" },
        },
      };

      let appliedCell;

      global.Excel.run = vi.fn(async (callback) => {
        const context = {
          workbook: {
            getSelectedRange: () => ({
              load: vi.fn(),
              address: "Sheet1!A1",
              rowCount: 1,
              columnCount: 1,
              getCell: () => {
                appliedCell = {
                  format: {
                    font: { bold: false, color: "" },
                    fill: { color: "" },
                    borders: {
                      getItem: () => ({
                        style: "",
                        color: "",
                        weight: "",
                      }),
                    },
                    protection: {
                      locked: false,
                      formulaHidden: false,
                    },
                    horizontalAlignment: "",
                    verticalAlignment: "",
                    wrapText: false,
                    indentLevel: 0,
                    readingOrder: "",
                    textOrientation: 0,
                  },
                  numberFormat: "",
                  numberFormatLocal: "",
                };
                return appliedCell;
              },
            }),
          },
          sync: vi.fn(),
        };
        await callback(context);
      });

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(
        JSON.stringify({ [styleName]: mockPreset }),
      );

      await loadCellStylePreset(styleName);

      expect(global.Excel.run).toHaveBeenCalled();
      expect(appliedCell.format.font.bold).toBe(true);
      expect(appliedCell.format.fill.color).toBe("red");
    });

    it("should display a warning if the preset does not exist", async () => {
      const styleName = "NonExistentStyle";

      OfficeRuntime.storage.getItem.mockResolvedValueOnce(JSON.stringify({}));

      await loadCellStylePreset(styleName);

      expect(cellCommonUtils.updateState).toHaveBeenCalledWith(
        "setMessageList",
        expect.objectContaining({
          type: "warning",
          title: "로드 오류",
          body: "프리셋을 불러오는데 실패하였습니다",
        }),
      );
    });
  });
});
