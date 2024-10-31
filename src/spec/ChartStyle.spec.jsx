import {
  render,
  screen,
  fireEvent,
  waitFor,
  act,
} from "@testing-library/react";

import ChartStyle from "../taskpane/components/Style/ChartStyle";
import { addPreset, deletePreset } from "../taskpane/utils/commonFuncs";
import {
  copyChartStyle,
  pasteChartStyle,
} from "../taskpane/utils/chartStyleFuncs";

vi.mock("../taskpane/utils/commonFuncs");
vi.mock("../taskpane/utils/cellStyleFuncs");
vi.mock("../taskpane/utils/chartStyleFuncs");

global.OfficeRuntime = {
  storage: {
    getItem: vi.fn().mockResolvedValue("{}"),
    setItem: vi.fn().mockResolvedValue(),
  },
};

describe("ChartStyle", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("should create a new chart style preset", async () => {
    addPreset.mockResolvedValue();

    render(<ChartStyle />);

    await act(async () => {
      fireEvent.click(screen.getByLabelText("plus"));
    });

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("chartStylePresets", "차트 서식1");
    });
  });

  it("should create a new chart style preset with incremented number", async () => {
    OfficeRuntime.storage.getItem.mockResolvedValue(
      JSON.stringify({ "차트 서식1": {}, "차트 서식2": {} }),
    );

    render(<ChartStyle />);

    await waitFor(() => {
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "chartStylePresets",
      );
    });

    fireEvent.click(screen.getByLabelText("plus"));

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("chartStylePresets", "차트 서식3");
    });
  });

  it("should handle chart style preset deletion", async () => {
    deletePreset.mockResolvedValue();

    const { rerender } = render(<ChartStyle />);

    act(() => {
      fireEvent.click(screen.getByLabelText("plus"));
    });

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("chartStylePresets", "차트 서식1");
    });

    rerender(<ChartStyle />);

    fireEvent.click(screen.getByLabelText("delete"));

    await waitFor(() => {
      expect(deletePreset).toHaveBeenCalledWith(
        "chartStylePresets",
        "차트 서식1",
      );
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "chartStylePresets",
      );
    });
  });

  it("should save the current chart style preset", async () => {
    render(<ChartStyle />);

    fireEvent.click(screen.getByLabelText("save"));

    await waitFor(() => {
      expect(copyChartStyle).toHaveBeenCalledWith("chartStylePresets", "");
    });
  });

  it("should load the selected chart style preset", async () => {
    render(<ChartStyle />);

    fireEvent.click(screen.getByText("적용"));

    await waitFor(() => {
      expect(pasteChartStyle).toHaveBeenCalledWith("chartStylePresets", "");
    });
  });
});
