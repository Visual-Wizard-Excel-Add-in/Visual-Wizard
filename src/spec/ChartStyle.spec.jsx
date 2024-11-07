import { render, screen, waitFor, act } from "@testing-library/react";
import { userEvent } from "@testing-library/user-event";
import { OfficeMockObject } from "office-addin-mock";

import ChartStyle from "../taskpane/components/Style/ChartStyle";
import { mockChartStyle, savedChartStyle } from "./test-constants";

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
      chartType: {
        ...mockChartStyle,
      },
      // _chartType: {
      //   load: vi.fn(),
      //   foramt: {
      //     fill: {
      //       setSolidColor: vi.fn((color) => (this.format.fill.color = color)),
      //       color: "",
      //     },
      //   },
      //   plotArea: {
      //     format: {
      //       fill: {
      //         color: "",
      //         setSolidColor: vi.fn(
      //           (color) =>
      //             (this.format.fill.plotArea.foramt.fill.color = color),
      //         ),
      //       },
      //     },
      //   },
      //   legend: {
      //     format: {
      //       fill: {
      //         color: "",
      //         setSolidColor: vi.fn(
      //           (color) =>
      //             (this.format.fill.plotArea.legend.fill.color = color),
      //         ),
      //       },
      //     },
      //   },
      //   series: {
      //     items: [
      //       {
      //         load: vi.fn(),
      //         format: {
      //           fill: {
      //             color: "",
      //             setSolidColor: vi.fn(
      //               (color) =>
      //                 (this.format.fill.plotArea.foramt.fill.color = color),
      //             ),
      //           },
      //         },
      //       },
      //     ],
      //   },
      // },
      getActiveChart: vi.fn(function () {
        // if ((await loadStorage()) === savedChartStyle) {
        //   return this._chartType;
        // }

        return this.chartType;
      }),
      getSelectedRange: vi.fn(function () {
        return this.range;
      }),
    },
  },
  run: vi.fn(async function (callback) {
    await callback(this.context);
  }),
};

async function loadStorage() {
  return JSON.parse(await OfficeRuntime.storage.getItem("chartStylePresets"));
}

vi.stubGlobal("Excel", new OfficeMockObject(mockData));
vi.stubGlobal("OfficeRuntime", OfficeRuntimeMock);

describe("ChartStyle", () => {
  it("프리셋 '+'버튼 클릭 시, 프리셋이 추가되어야 한다.", async () => {
    render(<ChartStyle />);

    await userEvent.click(screen.getByLabelText("add new preset"));

    expect(screen.getByText("차트 서식1")).toBeInTheDocument();
  });
  it("프리셋이 이미 존재할 경우, 프리셋을 추가할 경우 숫자가 1 증가해야합니다.", async () => {
    render(<ChartStyle />);

    await userEvent.click(screen.getByLabelText("add new preset"));

    expect(screen.getByText("차트 서식2")).toBeInTheDocument();
  });

  it("프리셋을 삭제할 수 있어야 한다.", async () => {
    render(<ChartStyle />);

    await userEvent.click(screen.getByLabelText("delete preset"));

    expect(await loadStorage()).toStrictEqual({ "차트 서식2": {} });
  });

  it("현재 서식을 저장할 수 있어야 한다.", async () => {
    render(<ChartStyle />);

    await userEvent.click(screen.getByLabelText("save button"));

    expect(await loadStorage()).toStrictEqual(savedChartStyle);
  });

  // it("should load the selected chart style preset", async () => {
  //   render(<ChartStyle />);

  //   await userEvent.click(screen.getByLabelText("paste button"));

  //   expect(mockData.context.workbook._chartType.format).toStrictEqual(
  //     savedChartStyle["차트 서식1"],
  //   );
  // });
});
