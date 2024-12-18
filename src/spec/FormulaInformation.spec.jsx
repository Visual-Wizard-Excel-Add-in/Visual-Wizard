import { render, screen } from "@testing-library/react";

import FormulaInformation from "../taskpane/components/Fomula/FormulaInfomation";
import createPubliceSlice from "../taskpane/store/createPublicSlice";
import FORMULA_EXPLANATION from "../taskpane/constants/formulaConstants";

vi.mock("../taskpane/utils/store", () => ({
  default: vi.fn(),
}));

describe("FormulaInformation", () => {
  beforeEach(() => {
    createPubliceSlice.mockReturnValue({
      cellFunctions: ["IF", "SUM", "AVERAGE"],
    });
  });

  it("Should show all functions in selected formula.", () => {
    render(<FormulaInformation />);

    expect(screen.getByText("IF")).toBeInTheDocument();
    expect(screen.getByText("SUM")).toBeInTheDocument();
    expect(screen.getByText("AVERAGE")).toBeInTheDocument();

    const ifExplanation = new RegExp(FORMULA_EXPLANATION["IF"].split("\n")[0]);
    const sumExplanation = new RegExp(
      FORMULA_EXPLANATION["SUM"].split("\n")[0],
    );
    const averageExplanation = new RegExp(
      FORMULA_EXPLANATION["AVERAGE"].split("\n")[0],
    );

    expect(screen.queryByText(ifExplanation)).toBeInTheDocument();
    expect(screen.queryByText(sumExplanation)).toBeInTheDocument();
    expect(screen.queryByText(averageExplanation)).toBeInTheDocument();
  });

  it("If selected cell has no Formula, should show notice message.", () => {
    createPubliceSlice.mockReturnValue({
      cellFunctions: [],
    });

    render(<FormulaInformation />);

    expect(
      screen.getByText("수식이 입력된 셀을 선택해주세요."),
    ).toBeInTheDocument();
  });
});
