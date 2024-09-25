import {
  render,
  screen,
  fireEvent,
  waitFor,
  act,
} from "@testing-library/react";

import ValidateTest from "../taskpane/components/Validate/ValidateTest";
import { getLastCellAddress } from "../taskpane/utils/commonFuncs";
import { detectErrorCell } from "../taskpane/utils/cellStyleFunc";

global.Excel = {
  run: vi.fn().mockImplementation(async (context) => {
    if (typeof context === "function") {
      const fakeContext = {
        workbook: {
          worksheets: {
            onSelectionChanged: {
              add: vi.fn(),
            },
            getActiveWorksheet: () => ({
              load: vi.fn(),
              id: "1",
            }),
          },
        },
        sync: vi.fn(),
      };
      return await context(fakeContext);
    }
    return Promise.resolve();
  }),
};

vi.mock("../taskpane/utils/cellStyleFunc");
vi.mock("../taskpane/utils/cellCommonUtils");

describe("ValidateTest", () => {
  it("When rendering the component, fetchLastcellAddress functions should called", async () => {
    const fetchLastCellAddress = vi.fn().mockResolvedValue("A1");
    getLastCellAddress.mockImplementation(fetchLastCellAddress);

    render(<ValidateTest />);

    await waitFor(() => {
      expect(fetchLastCellAddress).toHaveBeenCalled();
    });

    const container = screen.getByText("사용중인 마지막 셀 영역:").closest("p");
    expect(container).toBeInTheDocument();

    expect(container).toHaveTextContent("사용중인 마지막 셀 영역: A1");
  });

  it("When toggle the '에러 셀 검사' button, highlightError function should called", async () => {
    render(<ValidateTest />);

    const highlightButton = screen.getByLabelText("에러 셀 검사");

    await act(async () => {
      fireEvent.click(highlightButton);
    });

    expect(detectErrorCell).toHaveBeenCalledWith(true);
  });
});
