import {
  render,
  screen,
  fireEvent,
  waitFor,
  act,
} from "@testing-library/react";

import FormulaTest from "../taskpane/components/Validate/FormulaTest";
import useStore from "../taskpane/utils/store";
import {
  extractAddresses,
  evaluateTestFormula,
} from "../taskpane/utils/cellCommonUtils";
import {
  groupCellsIntoRanges,
  parseFormulaSteps,
} from "../taskpane/utils/cellFormulaFunc";
import { describe, it, vi } from "vitest";

vi.mock("../taskpane/utils/store", () => ({
  default: vi.fn(),
}));

vi.mock("../taskpane/utils/cellCommonUtils", () => ({
  extractAddresses: vi.fn(),
  evaluateTestFormula: vi.fn(),
}));

vi.mock("../taskpane/utils/cellFormulaFunc", () => ({
  parseFormulaSteps: vi.fn(),
  groupCellsIntoRanges: vi.fn(),
}));

describe("FormulaTest", () => {
  beforeEach(() => {
    useStore.mockReturnValue({
      cellFormula: "=SUM(A1:B1)",
      cellValue: "3",
      cellArguments: ["A1:B1"],
    });

    parseFormulaSteps.mockResolvedValue([{ address: "A1:B1" }]);

    extractAddresses.mockReturnValue(["A1", "B1"]);
    groupCellsIntoRanges.mockReturnValue(["A1:B1"]);
    evaluateTestFormula.mockResolvedValue("10");
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  it("should display the formula and cell value", async () => {
    await act(async () => {
      render(<FormulaTest />);
    });

    expect(screen.getByText("선택한 셀의 수식:")).toBeInTheDocument();
    expect(screen.getByText("=SUM(A1:B1)")).toBeInTheDocument();
    expect(screen.getByText("현재 결과:")).toBeInTheDocument();
    expect(screen.getByText("3")).toBeInTheDocument();
  });

  it("should display the arguments correctly", async () => {
    await act(async () => {
      render(<FormulaTest />);
    });

    const argElement = screen.getByText(/1\. 인자:A1:B1/);
    expect(argElement).toBeInTheDocument();

    const inputElement = screen.getByPlaceholderText("변경할 값이나 셀 주소");
    expect(inputElement).toBeInTheDocument();

    fireEvent.change(inputElement, { target: { value: "7" } });
    expect(inputElement.value).toBe("7");
  });

  it("should evaluate the formula and display the result when the execute button is clicked", async () => {
    await act(async () => {
      render(<FormulaTest />);
    });

    const inputElement = screen.getByPlaceholderText("변경할 값이나 셀 주소");
    fireEvent.change(inputElement, { target: { value: "7" } });

    const executeButton = screen.getByText("실행");
    await act(async () => {
      fireEvent.click(executeButton);
    });

    await waitFor(() => {
      expect(evaluateTestFormula).toHaveBeenCalledWith("=SUM(7)");
      expect(screen.getByText("테스트 결과: 10")).toBeInTheDocument();
    });
  });
});
