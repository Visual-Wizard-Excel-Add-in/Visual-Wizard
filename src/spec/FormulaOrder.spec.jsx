import { fireEvent, render, screen } from "@testing-library/react";

import FormulaOrder from "../taskpane/components/Fomula/FormulaOrder";
import useStore from "../taskpane/utils/store";
import { parseFormulaSteps } from "../taskpane/utils/cellFormulaFuncs";

vi.mock("../taskpane/utils/store", () => ({
  __esModule: true,
  default: vi.fn(),
}));

vi.mock("../taskpane/utils/cellFormulaFunc", () => ({
  parseFormulaSteps: vi.fn(),
}));

describe("FormulaOrder", () => {
  beforeEach(() => {
    useStore.mockReturnValue({
      cellFormula: "=IF(SUM(A1:B1)>AVERAGE(A1:C1),1,2)",
      formulaSteps: [
        { address: "A1", functionName: "AVERAGE" },
        { address: "A1", functionName: "SUM" },
        { address: "A1", functionName: "IF" },
      ],
      setFormulaSteps: vi.fn(),
    });
  });

  it("If select cell has no formula, Should show notice message.", () => {
    useStore.mockReturnValue({
      cellFormula: "",
      formulaSteps: [],
      setFormulaSteps: vi.fn(),
    });

    render(<FormulaOrder />);

    expect(
      screen.getByText("수식이 입력된 셀을 선택해주세요."),
    ).toBeInTheDocument();
  });

  it("FormulaOrder should show functions by ordered.", () => {
    render(<FormulaOrder />);

    const functions = screen.queryAllByText(/IF|SUM|AVERAGE/);
    expect(functions[0]).toHaveTextContent("AVERAGE");
    expect(functions[1]).toHaveTextContent("SUM");
    expect(functions[2]).toHaveTextContent("IF");
  });

  it("cellFormula is change, parseFormulaStpes should called.", () => {
    useStore.mockReturnValue({
      cellFormula: "",
      formulaSteps: [],
      setFormulaSteps: vi.fn(),
    });

    render(<FormulaOrder />);

    expect(parseFormulaSteps).toHaveBeenCalledTimes(1);
  });

  it("When click the cell's function names button, should show FormulaOrderDescription.", () => {
    const step = {
      address: "A1",
      condition: "SUM(A1) > AVERAGE(A1)",
      dependencies: [],
      falseValue: "2",
      formula: "IF(SUM(A1:B1)>AVERAGE(A1:C1),1,2)",
      functionName: "IF",
      trueValue: "1",
    };

    useStore.mockReturnValue({
      formulaSteps: [step],
      setFormulaSteps: vi.fn(),
    });

    render(<FormulaOrder />);

    const triggerButton = screen.getByRole("button", { name: "IF" });

    fireEvent.click(triggerButton);

    expect(screen.getByText("참조 셀: A1")).toBeInTheDocument();
    expect(screen.getByText("조건: SUM(A1) > AVERAGE(A1)")).toBeInTheDocument();
    expect(screen.getByText("참일 때: 1")).toBeInTheDocument();
    expect(screen.getByText("거짓일 때: 2")).toBeInTheDocument();
    expect(
      screen.getByText("식: IF(SUM(A1:B1)>AVERAGE(A1:C1),1,2)"),
    ).toBeInTheDocument();
  });
});
