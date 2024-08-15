import { render, screen, fireEvent } from "@testing-library/react";

import FormulaAttribute from "../taskpane/components/Fomula/FormulaAttribute";
import useStore from "../taskpane/utils/store";
import { highlightingCell } from "../taskpane/utils/cellStyleFunc";

vi.mock("../taskpane/utils/store", () => ({
  default: vi.fn(),
}));
vi.mock("../taskpane/utils/cellStyleFunc");
vi.mock("../taskpane/utils/cellFormulaFunc");
vi.mock("../taskpane/utils/cellCommonUtils");

describe("FormulaAttribute", () => {
  let mockSetIsCellHighlighting;
  let mockCellFunctions;

  beforeEach(() => {
    mockSetIsCellHighlighting = vi.fn();
    mockCellFunctions = [];

    useStore.mockReturnValue({
      isCellHighlighting: false,
      setIsCellHighlighting: mockSetIsCellHighlighting,
      cellFunctions: mockCellFunctions,
      cellArguments: ["A4(1)", "B4(2)"],
      cellAddress: "sheet1!C4",
    });
  });

  it("should enable the highlight button when cellFunctions is not empty", () => {
    mockCellFunctions = ["function1", "function2"];

    useStore.mockReturnValue({
      isCellHighlighting: false,
      setIsCellHighlighting: mockSetIsCellHighlighting,
      cellFunctions: mockCellFunctions,
      cellArguments: ["A4(1)", "B4(2)"],
      cellAddress: "sheet1!C4",
    });

    render(<FormulaAttribute />);

    const highlightButton = screen.getByRole("switch");

    expect(highlightButton).not.toBeDisabled();
  });

  it("When toggle the highlight button, should call highlightingCell function", () => {
    mockCellFunctions = ["function1", "function2"];

    useStore.mockReturnValue({
      isCellHighlighting: false,
      setIsCellHighlighting: mockSetIsCellHighlighting,
      cellFunctions: mockCellFunctions,
      cellArguments: ["A4(1)", "B4(2)"],
      cellAddress: "sheet1!C4",
    });

    render(<FormulaAttribute />);

    const highlightButton = screen.getByRole("switch");

    fireEvent.click(highlightButton);

    expect(highlightingCell).toHaveBeenCalledTimes(1);
  });
});
