import { vi, describe, it, expect, beforeEach } from "vitest";
import * as originalCellCommonUtils from "../taskpane/utils/commonFuncs";
import useStore from "../taskpane/utils/store";

const mockExcel = {
  run: vi.fn(),
  ErrorCodes: {
    itemNotFound: "ItemNotFound",
  },
};

const mockOfficeRuntime = {
  storage: {
    getItem: vi.fn(),
    setItem: vi.fn(),
  },
};

let cellCommonUtils;

vi.mock("../taskpane/utils/cellCommonUtils", async () => {
  const mockUpdateState = vi.fn();
  const actual = await vi.importActual("../taskpane/utils/cellCommonUtils");

  return {
    ...actual,
    updateState: mockUpdateState,
    getTargetCellValue: vi.fn(),
  };
});

const mockSetState = {
  setMessageList: vi.fn(),
  setCellAddress: vi.fn(),
  setCellValue: vi.fn(),
  setCellFormula: vi.fn(),
  setCellFunctions: vi.fn(),
  setCellArguments: vi.fn(),
};

vi.mock("../taskpane/utils/store", () => ({
  default: {
    getState: vi.fn(() => ({
      setMessageList: mockSetState.setMessageList,
      setCellAddress: mockSetState.setCellAddress,
      setCellValue: mockSetState.setCellValue,
      setCellFormula: mockSetState.setCellFormula,
      setCellFunctions: mockSetState.setCellFunctions,
      setCellArguments: mockSetState.setCellArguments,
    })),
  },
}));

describe("cellCommonUtils", () => {
  beforeEach(async () => {
    vi.resetAllMocks();

    global.Excel = mockExcel;
    global.OfficeRuntime = mockOfficeRuntime;

    cellCommonUtils = await vi.importMock("../taskpane/utils/cellCommonUtils");

    useStore.getState.mockReturnValue(mockSetState);
  });

  describe("splitCellAddress", () => {
    it("should correctly split a cell address", () => {
      expect(cellCommonUtils.splitCellAddress("A1")).toEqual(["A", 1]);
      expect(cellCommonUtils.splitCellAddress("$B$2")).toEqual(["B", 2]);
      expect(cellCommonUtils.splitCellAddress("AA100")).toEqual(["AA", 100]);
    });

    it("should throw an error for invalid cell addresses", () => {
      expect(() => cellCommonUtils.splitCellAddress("Invalid")).toThrow(
        "Invalid cell address: Invalid",
      );
    });
  });

  describe("extractAddresses", () => {
    it("should extract single cell addresses", () => {
      expect(cellCommonUtils.extractAddresses("A1")).toEqual(["A1"]);
      expect(cellCommonUtils.extractAddresses("Sheet1!B2")).toEqual(["B2"]);
    });

    it("should extract cell ranges", () => {
      expect(cellCommonUtils.extractAddresses("A1:B2")).toEqual([
        "A1",
        "A2",
        "B1",
        "B2",
      ]);
      expect(cellCommonUtils.extractAddresses("Sheet1!C3:D4")).toEqual([
        "C3",
        "C4",
        "D3",
        "D4",
      ]);
    });

    it("should handle multiple addresses and ranges", () => {
      expect(cellCommonUtils.extractAddresses("A1,B2:C3")).toEqual([
        "A1",
        "B2",
        "B3",
        "C2",
        "C3",
      ]);
    });
  });

  describe("getChartTypeInKorean", () => {
    it("should return the correct Korean chart type", () => {
      expect(cellCommonUtils.getChartTypeInKorean("ColumnClustered")).toBe(
        "묶은 세로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Pie")).toBe("원형 차트");
      expect(cellCommonUtils.getChartTypeInKorean("Line")).toBe(
        "꺾은선형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Invalid")).toBe(
        "유효하지 않음",
      );
      expect(cellCommonUtils.getChartTypeInKorean("ColumnStacked")).toBe(
        "누적 세로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("3DColumnClustered")).toBe(
        "3D 묶은 세로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("BarClustered")).toBe(
        "묶은 가로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("LineMarkers")).toBe(
        "표식이 있는 꺾은선형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Area")).toBe("영역형 차트");
      expect(cellCommonUtils.getChartTypeInKorean("Doughnut")).toBe(
        "도넛형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Histogram")).toBe(
        "히스토그램형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("XYScatter")).toBe(
        "분산형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Bubble")).toBe(
        "거품형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Waterfall")).toBe(
        "폭포 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("StockOHLC")).toBe(
        "시가-고가-저가-종가 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Surface")).toBe(
        "3차원 표면형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("Radar")).toBe("방사형 차트");
      expect(cellCommonUtils.getChartTypeInKorean("CylinderColClustered")).toBe(
        "원기둥 묶은 세로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("ConeColClustered")).toBe(
        "원뿔 묶은 세로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("PyramidColClustered")).toBe(
        "피라미드 묶은 세로 막대형 차트",
      );
      expect(cellCommonUtils.getChartTypeInKorean("RegionMap")).toBe(
        "지역 맵형 차트",
      );
    });

    it('should return "알 수 없는 차트 유형" for unknown chart types', () => {
      expect(cellCommonUtils.getChartTypeInKorean("UnknownType")).toBe(
        "알 수 없는 차트 유형",
      );
    });
  });

  describe("getChartTypeInEnglish", () => {
    it("should return the correct English chart type", () => {
      expect(
        cellCommonUtils.getChartTypeInEnglish("묶은 세로 막대형 차트"),
      ).toBe("ColumnClustered");
      expect(cellCommonUtils.getChartTypeInEnglish("원형 차트")).toBe("Pie");
      expect(cellCommonUtils.getChartTypeInEnglish("꺾은선형 차트")).toBe(
        "Line",
      );
    });

    it('should return "Unknown chart type" for unknown chart types', () => {
      expect(cellCommonUtils.getChartTypeInEnglish("알 수 없는 유형")).toBe(
        "Unknown chart type",
      );
    });
  });

  describe("evaluateTestFormula", () => {
    it("should evaluate a test formula correctly", async () => {
      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: vi.fn().mockReturnValue({
              load: vi.fn(),
              name: "Sheet1",
            }),
            getItem: vi.fn().mockReturnValue({
              delete: vi.fn(),
            }),
            add: vi.fn().mockReturnValue({
              getRange: vi.fn().mockReturnValue({
                formulas: null,
                load: vi.fn(),
                values: [["Test Result"]],
              }),
              delete: vi.fn(),
            }),
          },
        },
        sync: vi.fn(),
      };

      vi.mocked(Excel.run).mockImplementation((callback) =>
        callback(mockContext),
      );

      const result = await cellCommonUtils.evaluateTestFormula("SUM(A1:A5)");

      expect(result).toBe("Test Result");
    });

    it("should handle errors and return null", async () => {
      vi.mocked(Excel.run).mockRejectedValue(new Error("Test error"));

      const result =
        await cellCommonUtils.evaluateTestFormula("InvalidFormula");

      expect(result).toBeNull();
    });
  });

  describe("extractArgsAddress", () => {
    it("should extract address from regular cell references", () => {
      expect(cellCommonUtils.extractArgsAddress("A1")).toBe("A1");
      expect(cellCommonUtils.extractArgsAddress("$B$2")).toBe("B2");
      expect(cellCommonUtils.extractArgsAddress("Sheet1!C3")).toBe("C3");
    });

    it("should return null for non-cell references", () => {
      expect(cellCommonUtils.extractArgsAddress("NotAnAddress")).toBeNull();
      expect(cellCommonUtils.extractArgsAddress("123")).toBeNull();
    });
  });

  describe("getCellValue", () => {
    beforeEach(() => {
      vi.resetAllMocks();

      const mockUseStore = useStore;

      mockUseStore.getState.mockReturnValue(mockSetState);
    });

    it("should correctly fetch and update state with selected cell properties", async () => {
      const mockRange = {
        address: "Sheet1!A1",
        numberFormat: [["General"]],
        values: [["Test Value"]],
        formulas: [["=SUM(B1, B2)"]],
        load: vi.fn(),
      };

      const mockSheet = {
        getRange: vi.fn().mockReturnValue(mockRange),
      };

      const mockContext = {
        workbook: {
          getSelectedRange: vi.fn().mockReturnValue(mockRange),
          worksheets: {
            getActiveWorksheet: vi.fn().mockReturnValue(mockSheet),
          },
        },
        sync: vi.fn(),
      };

      vi.mocked(Excel.run).mockImplementation((callback) =>
        callback(mockContext),
      );

      vi.spyOn(cellCommonUtils, "extractFunctionsFromFormula").mockReturnValue([
        "SUM",
      ]);
      vi.spyOn(cellCommonUtils, "extractArgsFromFormula").mockResolvedValue([
        "B1(Test Value)",
        "B2(Test Value)",
      ]);

      await cellCommonUtils.getCellValue();

      expect(mockSetState.setCellAddress).toHaveBeenCalledWith("Sheet1!A1");
      expect(mockSetState.setCellValue).toHaveBeenCalledWith("Test Value");
      expect(mockSetState.setCellFormula).toHaveBeenCalledWith("=SUM(B1, B2)");
      expect(mockSetState.setCellFunctions).toHaveBeenCalledWith(["SUM"]);
      expect(mockSetState.setCellArguments).toHaveBeenCalledWith([
        "B1(Test Value)",
        "B2(Test Value)",
      ]);
    });

    it("should correctly handle date formatted cell values", async () => {
      const mockRange = {
        address: "Sheet1!A1",
        numberFormat: [["dd/mm/yy"]],
        values: [[44197]],
        formulas: [[""]],
        load: vi.fn(),
      };

      const mockContext = {
        workbook: {
          getSelectedRange: vi.fn().mockReturnValue(mockRange),
          worksheets: {
            getActiveWorksheet: vi.fn(),
          },
        },
        sync: vi.fn(),
      };

      mockExcel.run.mockImplementation((callback) => callback(mockContext));

      await cellCommonUtils.getCellValue();

      expect(mockSetState.setCellValue).toHaveBeenCalledWith("2021. 1. 1.");
    });
  });

  describe("getTargetCellValue", () => {
    it("should correctly fetch the value of a specified cell", async () => {
      const mockCell = {
        values: [["Target Value"]],
        numberFormat: [["General"]],
        load: vi.fn(),
      };

      const mockSheet = {
        getRange: vi.fn().mockReturnValue(mockCell),
      };

      const mockContext = {
        workbook: {
          worksheets: {
            getItem: vi.fn().mockReturnValue(mockSheet),
            getActiveWorksheet: vi.fn().mockReturnValue(mockSheet),
          },
        },
        sync: vi.fn(),
      };

      mockExcel.run.mockImplementation((callback) => callback(mockContext));

      const result = await cellCommonUtils.getTargetCellValue("Sheet1!B2");

      expect(result).toBe("Target Value");
    });

    it("should correctly handle date formatted target cell values", async () => {
      const mockCell = {
        values: [[44197]],
        numberFormat: [["dd/mm/yy"]],
        load: vi.fn(),
      };

      const mockSheet = {
        getRange: vi.fn().mockReturnValue(mockCell),
      };

      const mockContext = {
        workbook: {
          worksheets: {
            getItem: vi.fn().mockReturnValue(mockSheet),
            getActiveWorksheet: vi.fn().mockReturnValue(mockSheet),
          },
        },
        sync: vi.fn(),
      };

      mockExcel.run.mockImplementation((callback) => callback(mockContext));

      const result = await cellCommonUtils.getTargetCellValue("Sheet1!B2");

      expect(result).toBe("2021. 1. 1.");
    });
  });
});
