import {
  render,
  screen,
  fireEvent,
  waitFor,
  act,
} from "@testing-library/react";

import CellStyle from "../taskpane/components/Style/CellStyle";
import useStore from "../taskpane/utils/store";
import {
  saveCellStylePreset,
  loadCellStylePreset,
} from "../taskpane/utils/cellStyleFunc";
import { addPreset, deletePreset } from "../taskpane/utils/commonFuncs";

vi.mock("../taskpane/utils/store");
vi.mock("../taskpane/utils/cellCommonUtils");
vi.mock("../taskpane/utils/cellStyleFunc");

global.OfficeRuntime = {
  storage: {
    getItem: vi.fn().mockResolvedValue("{}"),
    setItem: vi.fn().mockResolvedValue(),
  },
};

describe("CellStyle", () => {
  let mockStore;

  beforeEach(() => {
    mockStore = {
      selectedStylePreset: "",
      setSelectedStylePreset: vi.fn(),
    };

    useStore.mockReturnValue(mockStore);
  });

  it("should create a new style preset", async () => {
    addPreset.mockResolvedValue();

    render(<CellStyle />);

    await act(async () => {
      fireEvent.click(screen.getByLabelText("plus"));
    });

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("cellStylePresets", "셀 서식1");
    });

    await waitFor(() => {
      expect(mockStore.setSelectedStylePreset).toHaveBeenCalledWith("셀 서식1");
    });
  });

  it("should create a new preset with incremented number if '셀 서식1' already exists", async () => {
    OfficeRuntime.storage.getItem.mockResolvedValue(
      JSON.stringify({ "셀 서식1": {}, "셀 서식2": {} }),
    );

    render(<CellStyle />);

    await waitFor(() => {
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "cellStylePresets",
      );
    });

    fireEvent.click(screen.getByLabelText("plus"));

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("cellStylePresets", "셀 서식3");
    });

    await waitFor(() => {
      expect(mockStore.setSelectedStylePreset).toHaveBeenCalledWith("셀 서식3");
    });
  });

  it("should handle preset deletion", async () => {
    deletePreset.mockResolvedValue();
    mockStore.selectedStylePreset = "셀 서식1";

    render(<CellStyle />);

    fireEvent.click(screen.getByLabelText("delete"));

    await waitFor(() => {
      expect(deletePreset).toHaveBeenCalledWith("cellStylePresets", "셀 서식1");
      expect(mockStore.setSelectedStylePreset).toHaveBeenCalledWith("");
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "cellStylePresets",
      );
    });
  });

  it("should save the current style preset", async () => {
    render(<CellStyle />);

    fireEvent.click(screen.getByLabelText("save"));

    await waitFor(() => {
      expect(saveCellStylePreset).toHaveBeenCalledWith(
        mockStore.selectedStylePreset,
      );
    });
  });

  it("should load the selected style preset", async () => {
    mockStore.selectedStylePreset = "셀 서식1";

    render(<CellStyle />);

    fireEvent.click(screen.getByText("적용"));

    await waitFor(() => {
      expect(loadCellStylePreset).toHaveBeenCalledWith("셀 서식1");
    });
  });
});
