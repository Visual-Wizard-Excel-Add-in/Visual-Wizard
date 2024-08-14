import { render, screen, waitFor, act } from "@testing-library/react";
import { describe, it, expect, beforeEach, vi } from "vitest";

import App from "../taskpane/components/App";
import useStore from "../taskpane/utils/store";

global.Office = {
  onReady: vi.fn(),
};

global.Excel = {
  run: async (callback) => {
    await callback({
      workbook: {
        worksheets: {
          onActivated: {
            add: vi.fn(),
          },
          getActiveWorksheet: () => ({
            load: vi.fn(),
            id: "1",
          }),
        },
      },
      sync: vi.fn(),
    });
  },
};

class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}

global.ResizeObserver = ResizeObserver;

vi.mock("../taskpane/utils/store");

describe("App", () => {
  let mockStore;

  beforeEach(() => {
    mockStore = {
      category: "Formula",
      setCategory: vi.fn((newCategory) => {
        mockStore.category = newCategory;
      }),
      setOpenTab: vi.fn(),
      openTab: [],
      sheetId: "1",
      setSheetId: vi.fn(),
      activeSheetId: vi.fn(),
      messageList: [],
    };

    useStore.mockImplementation(() => mockStore);
  });

  it("should render Formula features when the Formula category is selected", () => {
    render(<App />);

    expect(screen.getByText("정보")).toBeInTheDocument();
    expect(screen.getByText("참조")).toBeInTheDocument();
    expect(screen.getByText("순서")).toBeInTheDocument();
  });

  it("should render Style features when the Style category is set", async () => {
    await act(async () => {
      mockStore.setCategory("Style");
    });

    render(<App />);

    await waitFor(
      () => {
        expect(screen.getByText("셀 서식")).toBeInTheDocument();
        expect(screen.getByText("차트 서식")).toBeInTheDocument();
      },
      { timeout: 2000 },
    );
  });

  it("should render Macro features when the Macro category is set", async () => {
    await act(async () => {
      mockStore.setCategory("Macro");
    });

    render(<App />);

    await waitFor(
      () => {
        expect(screen.getByText("매크로 녹화")).toBeInTheDocument();
        expect(screen.getByText("매크로 설정")).toBeInTheDocument();
      },
      { timeout: 2000 },
    );
  });

  it("should render Validate features when the Validate category is set", async () => {
    await act(async () => {
      mockStore.setCategory("Validate");
    });

    render(<App />);

    await waitFor(
      () => {
        expect(screen.getByText("유효성 검사")).toBeInTheDocument();
        expect(screen.getByText("수식 테스트")).toBeInTheDocument();
      },
      { timeout: 2000 },
    );
  });

  it("should render Share features when the Share category is set", async () => {
    await act(async () => {
      mockStore.setCategory("Share");
    });

    render(<App />);

    await waitFor(
      () => {
        expect(screen.getByText("추출하기")).toBeInTheDocument();
      },
      { timeout: 2000 },
    );
  });
});
