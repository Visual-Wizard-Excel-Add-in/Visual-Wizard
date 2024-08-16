import {
  render,
  screen,
  fireEvent,
  waitFor,
  act,
} from "@testing-library/react";

import Extraction from "../taskpane/components/Share/Extraction";
import executeFunction from "../taskpane/utils/extractFileFunc";

vi.mock("../taskpane/utils/extractFileFunc", () => ({
  default: vi.fn(),
}));

class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}

global.ResizeObserver = ResizeObserver;

describe("Extraction", () => {
  it("When rendered Extraction, Should show NoticeBar front of anything", async () => {
    render(<Extraction />);

    const container = screen.getByText((content, element) => {
      const hasText = content.includes("추출하기 이용을 위해선");
      const isCorrectElement =
        element.className && typeof element.className === "string"
          ? element.className.includes("fui-MessageBarBody")
          : false;

      return hasText && isCorrectElement;
    });

    await waitFor(() => {
      expect(container).toHaveTextContent(
        /주의:추출하기 이용을 위해선.*먼저.*이곳.*을 방문해주세요!/,
      );
    });
  });

  it("When click the '저장' button, should call the executeFunction()", async () => {
    render(<Extraction />);

    const saveButton = screen.getByText("저장");

    act(() => {
      fireEvent.click(saveButton);
    });

    await waitFor(async () => {
      expect(executeFunction).toHaveBeenCalledWith("선택 영역");
    });
  });

  it("Should update the dataLocation when a new option is selected", async () => {
    render(<Extraction />);

    const dropdown = screen.getByText("선택 영역");

    fireEvent.click(dropdown);

    const newOption = screen.getByText("현재 시트");

    fireEvent.click(newOption);

    const saveButton = screen.getByText("저장");

    act(() => {
      fireEvent.click(saveButton);
    });

    await waitFor(() => {
      expect(executeFunction).toHaveBeenCalledWith("현재 시트");
    });
  });
});
