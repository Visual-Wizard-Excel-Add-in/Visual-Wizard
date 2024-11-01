import { render, screen, fireEvent, waitFor } from "@testing-library/react";

import MacroSetting from "../taskpane/components/Macro/MacroSetting";
import createPubliceSlice from "../taskpane/store/createPublicSlice";

global.OfficeRuntime = {
  storage: {
    getItem: vi.fn(),
    setItem: vi.fn(),
  },
};

vi.mock("../taskpane/utils/store", () => ({
  __esModule: true,
  default: vi.fn(),
}));

describe("MacroSetting", () => {
  beforeEach(() => {
    vi.clearAllMocks();

    createPubliceSlice.mockReturnValue({
      selectMacroPreset: "TestPreset",
    });

    global.OfficeRuntime.storage.getItem.mockResolvedValue(
      JSON.stringify({
        TestPreset: {
          actions: [
            {
              type: "WorksheetChanged",
              address: "A1",
              details: { value: "Test" },
            },
          ],
        },
      }),
    );
  });

  it("renders without crashing", async () => {
    render(<MacroSetting />);

    expect(
      await screen.findByText("선택한 프리셋: TestPreset"),
    ).toBeInTheDocument();
  });

  it("displays the correct number of actions", async () => {
    render(<MacroSetting />);

    const actions = await screen.findAllByText(/셀 내용 변경/);

    expect(actions).toHaveLength(1);
  });

  it("allows editing of action details", async () => {
    render(<MacroSetting />);

    const addressInput = await screen.findByPlaceholderText("A1");

    fireEvent.change(addressInput, { target: { value: "B1" } });

    expect(addressInput.value).toBe("B1");
  });

  it('calls applyChanges when "변경사항 적용" button is clicked', async () => {
    render(<MacroSetting />);

    const applyButton = await screen.findByText("변경사항 적용");

    fireEvent.click(applyButton);

    await waitFor(() => {
      expect(global.OfficeRuntime.storage.setItem).toHaveBeenCalled();
    });
  });

  it("handles different action types correctly", async () => {
    global.OfficeRuntime.storage.getItem.mockResolvedValue(
      JSON.stringify({
        TestPreset: {
          actions: [
            {
              type: "WorksheetChanged",
              address: "A1",
              details: { value: "Test" },
            },
            {
              type: "ChartAdded",
              chartId: "1",
              chartType: "ColumnClustered",
              dataRange: ["A1:B2"],
            },
          ],
        },
      }),
    );

    render(<MacroSetting />);

    expect(await screen.findByText(/셀 내용 변경/)).toBeInTheDocument();
    expect(await screen.findByText(/차트 추가/)).toBeInTheDocument();
  });

  it("displays an error message for unsupported action types", async () => {
    global.OfficeRuntime.storage.getItem.mockResolvedValue(
      JSON.stringify({
        TestPreset: {
          actions: [{ type: "UnsupportedType" }],
        },
      }),
    );

    render(<MacroSetting />);

    expect(
      await screen.findByText("지원하지 않는 형식의 기록입니다."),
    ).toBeInTheDocument();
  });
});
