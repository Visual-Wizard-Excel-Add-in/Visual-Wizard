import {
  render,
  fireEvent,
  screen,
  waitFor,
  act,
} from "@testing-library/react";

import MacroRecord from "../taskpane/components/Macro/MacroRecord";
import usePublicStore from "../taskpane/store/publicStore";
import {
  addPreset,
  deletePreset,
  updateState,
} from "../taskpane/utils/commonFuncs";
import { manageRecording, macroPlay } from "../taskpane/utils/macroFuncs";
import { vi } from "vitest";

vi.mock("../taskpane/utils/store");
vi.mock("../taskpane/utils/cellCommonUtils");
vi.mock("../taskpane/utils/macroFuncs");

global.OfficeRuntime = {
  storage: {
    getItem: vi.fn().mockResolvedValue("{}"),
    setItem: vi.fn().mockResolvedValue(),
  },
};

describe("MacroRecord Component", () => {
  let mockStore;

  beforeEach(() => {
    mockStore = {
      isRecording: false,
      setIsRecording: vi.fn(),
      selectMacroPreset: "매크로1",
      setSelectMacroPreset: vi.fn(),
    };

    usePublicStore.mockReturnValue(mockStore);
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  it("should create a new preset", async () => {
    addPreset.mockResolvedValue();

    render(<MacroRecord />);

    await act(async () => {
      fireEvent.click(screen.getByLabelText("plus"));
    });

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("allMacroPresets", "매크로1");
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "allMacroPresets",
      );
    });
  });

  it("should create a new preset with incremented number if '매크로1' already exists", async () => {
    OfficeRuntime.storage.getItem.mockResolvedValue(
      JSON.stringify({ 매크로1: {}, 매크로2: {} }),
    );

    render(<MacroRecord />);

    await waitFor(() => {
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "allMacroPresets",
      );
    });

    await act(async () => {
      fireEvent.click(screen.getByLabelText("plus"));
    });

    await waitFor(() => {
      expect(addPreset).toHaveBeenCalledWith("allMacroPresets", "매크로3");
    });
  });

  it("should handle preset deletion", async () => {
    deletePreset.mockResolvedValue();

    render(<MacroRecord />);

    await act(async () => {
      fireEvent.click(screen.getByLabelText("delete"));
    });

    await waitFor(() => {
      expect(deletePreset).toHaveBeenCalledWith("allMacroPresets", "매크로1");
      expect(mockStore.setSelectMacroPreset).toHaveBeenCalledWith("");
      expect(OfficeRuntime.storage.getItem).toHaveBeenCalledWith(
        "allMacroPresets",
      );
    });
  });

  it("should start and stop macro recording", async () => {
    render(<MacroRecord />);

    await act(async () => {
      fireEvent.click(screen.getByLabelText("record"));
    });

    expect(manageRecording).toHaveBeenCalledWith(true, "매크로1");
    expect(mockStore.setIsRecording).toHaveBeenCalledWith(true);
  });

  it("should show a warning if no preset is selected when recording", async () => {
    mockStore.selectMacroPreset = "";

    render(<MacroRecord />);

    await act(async () => {
      fireEvent.click(screen.getByLabelText("record"));
    });

    expect(updateState).toHaveBeenCalledWith("setMessageList", {
      type: "warning",
      title: "접근 오류:",
      body: "프리셋을 선택해주세요.",
    });
    expect(manageRecording).not.toHaveBeenCalled();
  });

  it("should play the macro when execute button is clicked", async () => {
    mockStore.selectMacroPreset = "매크로1";

    render(<MacroRecord />);

    await act(async () => {
      fireEvent.click(screen.getByText("실행"));
    });

    expect(macroPlay).toHaveBeenCalledWith("매크로1");
  });
});
