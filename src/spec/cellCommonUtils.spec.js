import { vi, describe, it, expect, beforeEach } from "vitest";
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

let commonFuncs;
let validateFuncs;

vi.mock("../taskpane/utils/commonFuncs", async () => {
  const mockUpdateState = vi.fn();
  const actual = await vi.importActual("../taskpane/utils/commonFuncs");

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

describe("commonFuncs", () => {
  beforeEach(async () => {
    vi.resetAllMocks();

    global.Excel = mockExcel;
    global.OfficeRuntime = mockOfficeRuntime;

    commonFuncs = await vi.importMock("../taskpane/utils/commonFuncs");
    validateFuncs = await vi.importMock("../taskpane/utils/validateFuncs");

    useStore.getState.mockReturnValue(mockSetState);
  });

  describe("splitCellAddress", () => {
    it("should correctly split a cell address", () => {
      expect(commonFuncs.splitCellAddress("A1")).toEqual(["A", 1]);
      expect(commonFuncs.splitCellAddress("$B$2")).toEqual(["B", 2]);
      expect(commonFuncs.splitCellAddress("AA100")).toEqual(["AA", 100]);
    });

    it("should throw an error for invalid cell addresses", () => {
      expect(() => commonFuncs.splitCellAddress("Invalid")).toThrow(
        "Invalid cell address: Invalid",
      );
    });
  });

  describe("extractReferenceCells", () => {
    it("should extract single cell addresses", () => {
      expect(commonFuncs.extractReferenceCells("A1")).toEqual(["A1"]);
      expect(commonFuncs.extractReferenceCells("Sheet1!B2")).toEqual(["B2"]);
    });

    it("should extract cell ranges", () => {
      expect(commonFuncs.extractReferenceCells("A1:B2")).toEqual([
        "A1",
        "A2",
        "B1",
        "B2",
      ]);
      expect(commonFuncs.extractReferenceCells("Sheet1!C3:D4")).toEqual([
        "C3",
        "C4",
        "D3",
        "D4",
      ]);
    });

    it("should handle multiple addresses and ranges", () => {
      expect(commonFuncs.extractReferenceCells("A1,B2:C3")).toEqual([
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
      expect(commonFuncs.getChartTypeInKorean("ColumnClustered")).toBe(
        "묶은 세로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("Pie")).toBe("원형 차트");
      expect(commonFuncs.getChartTypeInKorean("Line")).toBe("꺾은선형 차트");
      expect(commonFuncs.getChartTypeInKorean("Invalid")).toBe("유효하지 않음");
      expect(commonFuncs.getChartTypeInKorean("ColumnStacked")).toBe(
        "누적 세로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("3DColumnClustered")).toBe(
        "3D 묶은 세로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("BarClustered")).toBe(
        "묶은 가로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("LineMarkers")).toBe(
        "표식이 있는 꺾은선형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("Area")).toBe("영역형 차트");
      expect(commonFuncs.getChartTypeInKorean("Doughnut")).toBe("도넛형 차트");
      expect(commonFuncs.getChartTypeInKorean("Histogram")).toBe(
        "히스토그램형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("XYScatter")).toBe("분산형 차트");
      expect(commonFuncs.getChartTypeInKorean("Bubble")).toBe("거품형 차트");
      expect(commonFuncs.getChartTypeInKorean("Waterfall")).toBe("폭포 차트");
      expect(commonFuncs.getChartTypeInKorean("StockOHLC")).toBe(
        "시가-고가-저가-종가 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("Surface")).toBe(
        "3차원 표면형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("Radar")).toBe("방사형 차트");
      expect(commonFuncs.getChartTypeInKorean("CylinderColClustered")).toBe(
        "원기둥 묶은 세로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("ConeColClustered")).toBe(
        "원뿔 묶은 세로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("PyramidColClustered")).toBe(
        "피라미드 묶은 세로 막대형 차트",
      );
      expect(commonFuncs.getChartTypeInKorean("RegionMap")).toBe(
        "지역 맵형 차트",
      );
    });

    it('should return "알 수 없는 차트 유형" for unknown chart types', () => {
      expect(commonFuncs.getChartTypeInKorean("UnknownType")).toBe(
        "알 수 없는 차트 유형",
      );
    });
  });

  describe("getChartTypeInEnglish", () => {
    it("should return the correct English chart type", () => {
      expect(commonFuncs.getChartTypeInEnglish("묶은 세로 막대형 차트")).toBe(
        "ColumnClustered",
      );
      expect(commonFuncs.getChartTypeInEnglish("원형 차트")).toBe("Pie");
      expect(commonFuncs.getChartTypeInEnglish("꺾은선형 차트")).toBe("Line");
    });

    it('should return "Unknown chart type" for unknown chart types', () => {
      expect(commonFuncs.getChartTypeInEnglish("알 수 없는 유형")).toBe(
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

      const result = await validateFuncs.evaluateTestFormula("SUM(A1:A5)");

      expect(result).toBe("Test Result");
    });

    it("should handle errors and return null", async () => {
      vi.mocked(Excel.run).mockRejectedValue(new Error("Test error"));

      const result = await validateFuncs.evaluateTestFormula("InvalidFormula");

      expect(result).toBeNull();
    });
  });

  describe("extractArgsAddress", () => {
    it("should extract address from regular cell references", () => {
      expect(commonFuncs.extractArgsAddress("A1")).toBe("A1");
      expect(commonFuncs.extractArgsAddress("$B$2")).toBe("B2");
      expect(commonFuncs.extractArgsAddress("Sheet1!C3")).toBe("C3");
    });

    it("should return null for non-cell references", () => {
      expect(commonFuncs.extractArgsAddress("NotAnAddress")).toBeNull();
      expect(commonFuncs.extractArgsAddress("123")).toBeNull();
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

      vi.spyOn(commonFuncs, "extractFunctionsFromFormula").mockReturnValue([
        "SUM",
      ]);
      vi.spyOn(commonFuncs, "extractArgsFromFormula").mockResolvedValue([
        "B1(Test Value)",
        "B2(Test Value)",
      ]);

      await commonFuncs.getCellValue();

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

      await commonFuncs.getCellValue();

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

      const result = await commonFuncs.getTargetCellValue("Sheet1!B2");

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

      const result = await commonFuncs.getTargetCellValue("Sheet1!B2");

      expect(result).toBe("2021. 1. 1.");
    });
  });
});
