import { render, screen, act } from "@testing-library/react";
import { userEvent } from "@testing-library/user-event";
import { OfficeMockObject } from "office-addin-mock";

import App from "../taskpane/components/App";

const mockData = {
  context: {
    workbook: {
      worksheets: {
        id: "sheetName",
        getActiveWorksheet: function () {
          return {
            load: function () {},
          };
        },
        onActivated: {
          add: function () {},
        },
      },
    },
  },
  run: async function (callback) {
    await callback(this.context);
  },
  BorderWeight: {
    thick: "thick",
  },
  BorderLineStyle: {
    continuous: "continuous",
  },
};

global.Office = {
  onReady: vi.fn(),
};

global.Excel = new OfficeMockObject(mockData);

describe("App", () => {
  it("가장 초기 화면에서는 수식 카테고리의 기능들이 표시돼야한다.", async () => {
    await act(async () => render(<App />));

    expect(screen.getByText("정보")).toBeInTheDocument();
    expect(screen.getByText("참조")).toBeInTheDocument();
    expect(screen.getByText("순서")).toBeInTheDocument();
  });

  it("서식 카테고리를 클릭할 경우, 서식 기능들이 표시돼야한다", async () => {
    await act(async () => {
      render(<App />);
    });

    const styleTab = screen.getAllByText("서식", { exact: true });

    await userEvent.click(styleTab[0]);

    expect(screen.getByText("셀 서식")).toBeInTheDocument();
    expect(screen.getByText("차트 서식")).toBeInTheDocument();
  });

  it("매크로 카테고리를 클릭할 경우, 매크로 기능들이 표시돼야한다", async () => {
    await act(async () => {
      render(<App />);
    });

    const styleTab = screen.getAllByText("매크로", { exact: true });

    await userEvent.click(styleTab[0]);

    expect(screen.getByText("매크로 녹화")).toBeInTheDocument();
    expect(screen.getByText("매크로 설정")).toBeInTheDocument();
  });

  it("유효성 카테고리를 클릭할 경우, 유효성 기능들이 표시돼야한다", async () => {
    await act(async () => {
      render(<App />);
    });

    const styleTab = screen.getAllByText("유효성", { exact: true });

    await userEvent.click(styleTab[0]);

    expect(screen.getByText("유효성 검사")).toBeInTheDocument();
    expect(screen.getByText("수식 테스트")).toBeInTheDocument();
  });

  it("공유하기 카테고리를 클릭할 경우, 공유하기 기능들이 표시돼야한다", async () => {
    await act(async () => render(<App />));

    const styleTab = screen.getAllByText("공유하기", { exact: true });

    await userEvent.click(styleTab[0]);

    expect(screen.getByText("추출하기")).toBeInTheDocument();
  });
});
