import { render, screen } from "@testing-library/react";
import { userEvent } from "@testing-library/user-event";
import { OfficeMockObject } from "office-addin-mock";

import CellStyle from "../taskpane/components/Style/CellStyle";
import { mockStyle, savedStyle } from "./test-constants";

const OfficeRuntimeMock = {
  storage: {
    _storage: {},
    getItem: vi.fn(async (key) =>
      JSON.stringify(OfficeRuntimeMock.storage._storage[key] || {}),
    ),
    setItem: vi.fn(async (key, value) => {
      OfficeRuntimeMock.storage._storage[key] = JSON.parse(value);
    }),
  },
};

const mockData = {
  context: {
    workbook: {
      range: {
        address: "Sheet!A1",
        format: {},
        getCellProperties: vi.fn(function () {
          return {
            value: mockStyle,
          };
        }),
        setCellProperties: vi.fn(
          (source) => (mockData.context.workbook.range.format = source),
        ),
        copyFrom: vi.fn(),
      },
      getSelectedRange: function () {
        return this.range;
      },
      worksheets: {
        add: vi.fn(() => mockData.context.workbook.worksheet),
        getItemOrNullObject: vi.fn((sheetName) => {
          if ("StyleSheet" === sheetName) {
            return {
              ...mockData.context.workbook.worksheet,
              isNullObject: false,
            };
          }

          return { isNullObject: true };
        }),
      },
      worksheet: {
        delete: vi.fn(),
        setCellRange: vi.fn(),
        getRange: vi.fn(() => mockData.context.workbook.range),
      },
    },
  },
  run: vi.fn(async function (callback) {
    await callback(this.context);
  }),
};

async function loadStorage() {
  return JSON.parse(await OfficeRuntime.storage.getItem("cellStylePresets"));
}

vi.stubGlobal("Excel", new OfficeMockObject(mockData));
vi.stubGlobal("OfficeRuntime", OfficeRuntimeMock);

describe("CellStyle", () => {
  it("프리셋 '+'버튼 클릭 시, 프리셋이 추가되어야 한다.", async () => {
    render(<CellStyle />);

    await userEvent.click(screen.getByLabelText("add new preset"));

    expect(screen.getByText("셀 서식1")).toBeInTheDocument();
  });

  it("프리셋이 이미 존재할 경우, 프리셋을 추가할 경우 숫자가 1 증가해야합니다.", async () => {
    render(<CellStyle />);

    await userEvent.click(screen.getByLabelText("add new preset"));

    expect(screen.getByText("셀 서식2")).toBeInTheDocument();
  });

  it("프리셋을 삭제할 수 있어야한다", async () => {
    render(<CellStyle />);

    await userEvent.click(screen.getByLabelText("delete preset"));

    expect(await loadStorage()).toStrictEqual({ "셀 서식2": {} });
  });

  it("현재 서식을 저장할 수 있어야 한다.", async () => {
    render(<CellStyle />);

    await userEvent.click(screen.getByLabelText("add new preset"));
    await userEvent.click(screen.getByLabelText("save button"));

    expect(await loadStorage()).toStrictEqual(savedStyle);
  });

  it("선택한 셀 서식 적용 시, 올바른 서식이 불러와져야 한다", async () => {
    render(<CellStyle />);

    await userEvent.click(screen.getByLabelText("paste button"));

    expect(mockData.context.workbook.range.format).toStrictEqual(
      savedStyle["셀 서식1"][0],
    );
  });
});
