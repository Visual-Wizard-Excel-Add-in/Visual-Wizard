import { describe, it, expect, vi } from "vitest";
import {
  parseFormulaSteps,
  parseNestedFormula,
  processFunction,
} from "../taskpane/utils/cellFormulaFunc";

global.Excel = {
  run: vi.fn((callback) => {
    const context = {
      workbook: {
        getSelectedRange: vi.fn(() => ({
          load: vi.fn(),
          sync: vi.fn(),
          formulas: [[""]],
        })),
        worksheets: {
          getActiveWorksheet: vi.fn(() => ({
            getRange: vi.fn(() => ({
              load: vi.fn(),
              sync: vi.fn(),
              values: [[123]],
              numberFormat: [["General"]],
            })),
          })),
        },
      },
    };
    return callback(context);
  }),
};

describe("cellFormulaFunc Tests", () => {
  it("should return empty array if no formula in range", async () => {
    const result = await parseFormulaSteps();
    expect(result).toEqual([]);
  });

  it("should parse nested formula correctly", async () => {
    const context = {
      workbook: {
        worksheets: {
          getActiveWorksheet: vi.fn(() => ({
            getRange: vi.fn(() => ({
              load: vi.fn(),
              sync: vi.fn(),
              values: [[123]],
              numberFormat: [["General"]],
            })),
          })),
        },
      },
    };
    const formula = "SUM(A1:A10)";

    const result = await parseNestedFormula(context, formula);
    expect(result).toBeInstanceOf(Array);
    expect(result.length).toBeGreaterThan(0);
  });

  it("should process a simple function correctly", async () => {
    const context = {
      workbook: {
        worksheets: {
          getActiveWorksheet: vi.fn(() => ({
            getRange: vi.fn(() => ({
              load: vi.fn(),
              sync: vi.fn(),
              values: [[123]],
              numberFormat: [["General"]],
            })),
          })),
        },
      },
    };
    const funcName = "SUM";
    const args = "A1:A10";

    const result = await processFunction(context, funcName, args);
    expect(result.functionName).toBe("SUM");
    expect(result.address).toContain("A1:A10");
  });
});
