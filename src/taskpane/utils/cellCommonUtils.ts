import useStore from "./store";

function updateState(
  setStateFunc: keyof ReturnType<typeof useStore.getState>,
  newValue: string | number | boolean | string[] | MessageType | null,
) {
  const state = useStore.getState();

  if (typeof state[setStateFunc] === "function") {
    (state[setStateFunc] as (value: any) => void)(newValue);
  } else {
    console.error(`${setStateFunc} is not a function`);
  }
}

function splitCellAddress(address: string): [string, number] {
  const match = address.match(/\$?([A-Z]+)\$?([0-9]+)/);

  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  return [match[1], parseInt(match[2], 10)];
}

function extractAddresses(arg: string): string[] {
  const argAddresses: string[] = [];
  const argRegex = /((?:[^!]+!)?\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)/g;
  let match;

  while ((match = argRegex.exec(arg)) !== null) {
    const parts = match[0].split("!");
    const normalizedAddress = parts[parts.length - 1].replace(/\$/g, "");

    if (normalizedAddress.includes(":")) {
      const [startCell, endCell] = normalizedAddress.split(":");
      const cellsInRange = getCellsInRange(startCell, endCell);

      argAddresses.push(...cellsInRange);
    } else {
      argAddresses.push(normalizedAddress);
    }
  }

  return argAddresses;
}

function extractArgsAddress(cellArgument: string): string | null {
  const cleanedArgument = cellArgument.replace(/\$/g, "");
  const match = cleanedArgument.match(/([A-Z]+\d+)/);

  return match ? match[1] : null;
}

async function getCellValue(): Promise<void> {
  updateState("setCellFunctions", "");

  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const range = context.workbook.getSelectedRange();

      range.load([
        "address",
        "format/font",
        "formulas",
        "values",
        "numberFormat",
      ]);
      await context.sync();

      const selectedCellAddress = range.address;
      const numberFormat: string = range.numberFormat[0][0];
      let selectedCellValue: string | number | null = range.values[0][0];
      let formula: string = range.formulas[0][0];

      if (
        !range ||
        !range.values ||
        !range.values[0] ||
        range.values[0][0] === undefined
      ) {
        return;
      }

      if (typeof formula !== "string") {
        formula = "";
      }

      const formulaFunctions = extractFunctionsFromFormula(formula);
      const formulaArgs = await extractArgsFromFormula(formula);

      if (
        typeof selectedCellValue === "number" &&
        numberFormat &&
        numberFormat.includes("yy")
      ) {
        selectedCellValue = new Date(
          (selectedCellValue - 25569) * 86400 * 1000,
        ).toLocaleDateString();
      }

      updateState("setCellAddress", selectedCellAddress);
      updateState("setCellValue", selectedCellValue);
      updateState("setCellFormula", range.formulas[0][0]);
      updateState("setCellFunctions", formulaFunctions);
      updateState("setCellArguments", formulaArgs);

      await context.sync();
    });
  } catch (e: unknown) {
    if (e instanceof Error) {
      throw new Error(e.message);
    }
  }
}

async function getTargetCellValue(
  targetCell: string,
): Promise<string | number | null> {
  return Excel.run(async (context) => {
    const parts = targetCell.split("!");
    const sheetName = parts.length > 1 ? parts[0] : undefined;
    const normalizedAddress =
      parts.length > 1
        ? parts[1].replace(/\$/g, "")
        : parts[0].replace(/\$/g, "");

    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(normalizedAddress);

    cell.load(["values", "numberFormat"]);
    await context.sync();

    if (!cell.values || !cell.values[0] || cell.values[0][0] === undefined) {
      return null;
    }

    const numberFormat: string = cell.numberFormat[0][0];
    let targetCellValue: string | number | null = cell.values[0][0];

    if (
      numberFormat &&
      numberFormat.includes("yy") &&
      typeof targetCellValue === "number"
    ) {
      targetCellValue = new Date(
        (targetCellValue - 25569) * 86400 * 1000,
      ).toLocaleDateString();
    }

    return targetCellValue;
  });
}

async function extractArgsFromFormula(formula: string): Promise<string[]> {
  const argSet: Set<string> = new Set();
  const argCellRegex =
    /([A-Z]+[0-9]+|\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)/g;
  const results: string[] = [];
  const matches: string[] | null = formula.match(argCellRegex);

  if (matches) {
    for (const matchedArg of matches) {
      if (matchedArg.includes(":")) {
        const [startCell, endCell] = matchedArg.split(":");
        const cellsInRange = getCellsInRange(startCell, endCell);

        for (const cell of cellsInRange) {
          if (!argSet.has(cell)) {
            const value = await getTargetCellValue(cell);

            argSet.add(cell);
            results.push(`${cell}(${value})`);
          }
        }
      } else if (!argSet.has(matchedArg)) {
        const value = await getTargetCellValue(matchedArg);

        argSet.add(matchedArg);
        results.push(`${matchedArg}(${value})`);
      }
    }
  }

  return results;
}

function extractFunctionsFromFormula(formula: string): string[] {
  const functionNames: string[] = [];
  const regex = /([A-Z]+)\(/g;
  let match: RegExpExecArray | null;

  while ((match = regex.exec(formula)) !== null) {
    if (!functionNames.includes(match[1])) {
      functionNames.push(match[1]);
    }
  }

  return functionNames;
}

function getCellsInRange(startCell: string, endCell: string): string[] {
  const cells: string[] = [];
  const startColumn = startCell.match(/[A-Z]+/)?.[0] ?? null;
  const startRow = parseInt(startCell.match(/[0-9]+/)?.[0] ?? "0", 10);
  const endColumn = endCell.match(/[A-Z]+/)?.[0] ?? null;
  const endRow = parseInt(endCell.match(/[0-9]+/)?.[0] ?? "0", 10);
  let currentColumn = "";

  if (typeof startColumn === "string") {
    currentColumn = startColumn;
  }

  if (!startColumn || !startRow || !endColumn || !endRow) {
    throw new Error("Invalid cell address format");
  }

  while (currentColumn <= endColumn) {
    for (let row = startRow; row <= endRow; row += 1) {
      cells.push(`${currentColumn}${row}`);
    }

    if (currentColumn === endColumn) {
      break;
    }

    currentColumn = nextColumn(currentColumn);
  }

  return cells;
}

function nextColumn(col: string): string {
  if (col.length === 1) {
    return String.fromCharCode(col.charCodeAt(0) + 1);
  }

  let lastChar = col.slice(-1);
  let restChars = col.slice(0, -1);

  if (lastChar === "Z") {
    restChars = nextColumn(restChars);
    lastChar = "A";
  } else {
    lastChar = String.fromCharCode(lastChar.charCodeAt(0) + 1);
  }

  return restChars + lastChar;
}

async function registerSelectionChange(
  sheetId: string,
  func: () => Promise<void>,
) {
  await Excel.run(async (context) => {
    const { workbook } = context;
    const sheet = workbook.worksheets.getItem(sheetId);

    sheet.onSelectionChanged.add(func);
    await context.sync();
  });
}

async function activeSheetId(sheetId: string): Promise<void> {
  await Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();

    activeSheet.load("name");
    await context.sync();

    const activatedSheetId = activeSheet.name;

    if (activatedSheetId !== sheetId) {
      updateState("setSheetId", activatedSheetId);
      await registerSelectionChange(activatedSheetId, getCellValue);
    }
  });
}

async function addPreset(
  presetCategory: string,
  presetName: string,
): Promise<void> {
  let savePreset: string | { [key: string]: unknown } | null =
    await OfficeRuntime.storage.getItem(presetCategory);

  if (!savePreset) {
    savePreset = {};
  } else if (typeof savePreset === "string") {
    savePreset = JSON.parse(savePreset);
  }

  if (typeof savePreset === "object" && savePreset !== null) {
    savePreset[presetName] = {};

    await OfficeRuntime.storage.setItem(
      presetCategory,
      JSON.stringify(savePreset),
    );
  }
}

async function deletePreset(
  presetCategory: string,
  presetName: string,
): Promise<void> {
  let currentPresets: string | { [key: string]: unknown } =
    await OfficeRuntime.storage.getItem(presetCategory);

  if (currentPresets && typeof currentPresets === "string") {
    currentPresets = JSON.parse(currentPresets);

    if (typeof currentPresets === "object") {
      delete currentPresets[presetName];
    }

    await OfficeRuntime.storage.setItem(
      presetCategory,
      JSON.stringify(currentPresets),
    );
  }
}

async function getLastCellAddress(): Promise<string | null> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();

    usedRange.load(["values"]);
    await context.sync();

    let lastRowIndex = -1;
    let lastColIndex = -1;

    usedRange.values.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        if (value !== null && value !== "") {
          if (rowIndex > lastRowIndex) lastRowIndex = rowIndex;

          if (colIndex > lastColIndex) lastColIndex = colIndex;
        }
      });
    });

    if (lastRowIndex === -1 || lastColIndex === -1) {
      return null;
    }

    const lastCell = usedRange.getCell(lastRowIndex, lastColIndex);

    lastCell.load("address");
    await context.sync();

    const lastCellAddress = lastCell.address.split("!")[1];

    return lastCellAddress;
  });
}

function getChartTypeInKorean(chartTypeEnglish: string): string {
  switch (chartTypeEnglish) {
    case "Invalid":
      return "유효하지 않음";

    case "ColumnClustered":
      return "묶은 세로 막대형 차트";

    case "ColumnStacked":
      return "누적 세로 막대형 차트";

    case "ColumnStacked100":
      return "100% 기준 누적 세로 막대형 차트";

    case "3DColumnClustered":
      return "3D 묶은 세로 막대형 차트";

    case "3DColumnStacked":
      return "3D 누적 세로 막대형 차트";

    case "3DColumnStacked100":
      return "3D 100% 기준 누적 세로 막대형 차트";

    case "3DColumn":
      return "3차원 세로 막대형 차트";

    case "BarClustered":
      return "묶은 가로 막대형 차트";

    case "BarStacked":
      return "누적 가로 막대형 차트";

    case "BarStacked100":
      return "100% 기준 누적 가로 막대형 차트";

    case "3DBarClustered":
      return "3차원 묶은 가로 막대형 차트";

    case "3DBarStacked":
      return "3차원 누적 가로 막대형 차트";

    case "3DBarStacked100":
      return "3차원 100% 기준 누적 가로 막대형 차트";

    case "Line":
      return "꺾은선형 차트";

    case "LineStacked":
      return "누적 꺾은선형 차트";

    case "LineStacked100":
      return "100% 기준 누적 꺾은선형 차트";

    case "LineMarkers":
      return "표식이 있는 꺾은선형 차트";

    case "LineMarkersStacked":
      return "표식이 있는 누적 꺾은선형 차트";

    case "LineMarkersStacked100":
      return "표식이 있는 100% 기준 누적 꺾은선형 차트";

    case "3DLine":
      return "3D 꺾은선형 차트";

    case "Area":
      return "영역형 차트";

    case "AreaStacked":
      return "누적 영역형 차트";

    case "AreaStacked100":
      return "100% 기준 누적 영역형 차트";

    case "3DArea":
      return "3D 영역형 차트";

    case "3DAreaStacked":
      return "3D 누적 영역형 차트";

    case "3DAreaStacked100":
      return "3D 100% 기준 누적 영역형 차트";

    case "Pie":
      return "원형 차트";

    case "PieOfPie":
      return "원형 대 원형 차트";

    case "BarOfPie":
      return "원형 대 가로 막대형 차트";

    case "3DPie":
      return "3차원 원형 차트";

    case "PieExploded":
      return "분리형 원형 차트";

    case "3DPieExploded":
      return "3D 분리형 원형 차트";

    case "Doughnut":
      return "도넛형 차트";

    case "DoughnutExploded":
      return "분리형 도넛형 차트";

    case "Treemap":
      return "트리맵 차트";

    case "Sunburst":
      return "선버스트 차트";

    case "Histogram":
      return "히스토그램형 차트";

    case "Pareto":
      return "파레토 차트";

    case "BoxWhisker":
      return "상자 수염 차트";

    case "XYScatter":
      return "분산형 차트";

    case "XYScatterSmooth":
      return "곡선 및 표식이 있는 분산형 차트";

    case "XYScatterSmoothNoMarkers":
      return "곡선 및 표식이 없는 분산형 차트";

    case "XYScatterLines":
      return "직선 및 표식이 있는 분산형 차트";

    case "XYScatterLinesNoMarkers":
      return "직선 및 표식이 없는 분산형 차트";

    case "Bubble":
      return "거품형 차트";

    case "Bubble3DEffect":
      return "3차원 거품형 차트";

    case "Waterfall":
      return "폭포 차트";

    case "Funnel":
      return "깔때기형 차트";

    case "StockHLC":
      return "고가-저가-종가 차트";

    case "StockOHLC":
      return "시가-고가-저가-종가 차트";

    case "StockVHLC":
      return "거래량-고가-저가-종가 차트";

    case "StockVOHLC":
      return "거래량-시가-고가-저가-종가 차트";

    case "Surface":
      return "3차원 표면형 차트";

    case "SurfaceWireframe":
      return "3차원 표면형 차트(골격형)";

    case "SurfaceTopView":
      return "표면형 차트(조감도)";

    case "SurfaceTopViewWireframe":
      return "표면형 차트(골격형 조감도)";

    case "Radar":
      return "방사형 차트";

    case "RadarMarkers":
      return "표식이 있는 방사형 차트";

    case "RadarFilled":
      return "채워진 방사형 차트";

    case "CylinderColClustered":
      return "원기둥 묶은 세로 막대형 차트";

    case "CylinderColStacked":
      return "원기둥 누적 세로 막대형 차트";

    case "CylinderColStacked100":
      return "원기둥 100% 기준 누적 세로 막대형 차트";

    case "CylinderBarClustered":
      return "원기둥 묶은 가로 막대형 차트";

    case "CylinderBarStacked":
      return "원기둥 누적 가로 막대형 차트";

    case "CylinderBarStacked100":
      return "원기둥 100% 기준 누적 가로 막대형 차트";

    case "CylinderCol":
      return "원기둥 세로 막대형 차트";

    case "ConeColClustered":
      return "원뿔 묶은 세로 막대형 차트";

    case "ConeColStacked":
      return "원뿔 누적 세로 막대형 차트";

    case "ConeColStacked100":
      return "원뿔 100% 기준 누적 세로 막대형 차트";

    case "ConeBarClustered":
      return "원뿔 묶은 가로 막대형 차트";

    case "ConeBarStacked":
      return "원뿔 누적 가로 막대형 차트";

    case "ConeBarStacked100":
      return "원뿔 100% 기준 누적 가로 막대형 차트";

    case "ConeCol":
      return "원뿔 세로 막대형 차트";

    case "PyramidColClustered":
      return "피라미드 묶은 세로 막대형 차트";

    case "PyramidColStacked":
      return "피라미드 누적 세로 막대형 차트";

    case "PyramidColStacked100":
      return "피라미드 100% 기준 누적 세로 막대형 차트";

    case "PyramidBarClustered":
      return "피라미드 묶은 가로 막대형 차트";

    case "PyramidBarStacked":
      return "피라미드 누적 가로 막대형 차트";

    case "PyramidBarStacked100":
      return "피라미드 100% 기준 누적 가로 막대형 차트";

    case "PyramidCol":
      return "피라미드 세로 막대형 차트";

    case "RegionMap":
      return "지역 맵형 차트";

    default:
      return "알 수 없는 차트 유형";
  }
}

function getChartTypeInEnglish(chartTypeKorean: string): string {
  switch (chartTypeKorean) {
    case "유효하지 않음":
      return "Invalid";

    case "묶은 세로 막대형 차트":
      return "ColumnClustered";

    case "누적 세로 막대형 차트":
      return "ColumnStacked";

    case "100% 기준 누적 세로 막대형 차트":
      return "ColumnStacked100";

    case "3D 묶은 세로 막대형 차트":
      return "3DColumnClustered";

    case "3D 누적 세로 막대형 차트":
      return "3DColumnStacked";

    case "3D 100% 기준 누적 세로 막대형 차트":
      return "3DColumnStacked100";

    case "3차원 세로 막대형 차트":
      return "3DColumn";

    case "묶은 가로 막대형 차트":
      return "BarClustered";

    case "누적 가로 막대형 차트":
      return "BarStacked";

    case "100% 기준 누적 가로 막대형 차트":
      return "BarStacked100";

    case "3차원 묶은 가로 막대형 차트":
      return "3DBarClustered";

    case "3차원 누적 가로 막대형 차트":
      return "3DBarStacked";

    case "3차원 100% 기준 누적 가로 막대형 차트":
      return "3DBarStacked100";

    case "꺾은선형 차트":
      return "Line";

    case "누적 꺾은선형 차트":
      return "LineStacked";

    case "100% 기준 누적 꺾은선형 차트":
      return "LineStacked100";

    case "표식이 있는 꺾은선형 차트":
      return "LineMarkers";

    case "누적 꺾은선형 표식 차트":
      return "LineMarkersStacked";

    case "100% 기준 누적 꺾은선형 표식 차트":
      return "LineMarkersStacked100";

    case "3D 꺾은선형 차트":
      return "3DLine";

    case "영역형 차트":
      return "Area";

    case "누적 영역형 차트":
      return "AreaStacked";

    case "100% 기준 누적 영역형 차트":
      return "AreaStacked100";

    case "3D 영역형 차트":
      return "3DArea";

    case "3D 누적 영역형 차트":
      return "3DAreaStacked";

    case "3D 100% 기준 누적 영역형 차트":
      return "3DAreaStacked100";

    case "원형 차트":
      return "Pie";

    case "원형 대 원형 차트":
      return "PieOfPie";

    case "원형 대 가로 막대형 차트":
      return "BarOfPie";

    case "3차원 원형 차트":
      return "3DPie";

    case "분리형 원형 차트":
      return "PieExploded";

    case "3D 분리형 원형 차트":
      return "3DPieExploded";

    case "도넛형 차트":
      return "Doughnut";

    case "분리형 도넛형 차트":
      return "DoughnutExploded";

    case "트리맵 차트":
      return "Treemap";

    case "선버스트 차트":
      return "Sunburst";

    case "히스토그램형 차트":
      return "Histogram";

    case "파레토 차트":
      return "Pareto";

    case "상자 수염 차트":
      return "BoxWhisker";

    case "분산형 차트":
      return "XYScatter";

    case "곡선 및 표식이 있는 분산형 차트":
      return "XYScatterSmooth";

    case "곡선 및 표식이 없는 분산형 차트":
      return "XYScatterSmoothNoMarkers";

    case "직선 및 표식이 있는 분산형 차트":
      return "XYScatterLines";

    case "직선 및 표식이 없는 분산형 차트":
      return "XYScatterLinesNoMarkers";

    case "거품형 차트":
      return "Bubble";

    case "3차원 거품형 차트":
      return "Bubble3DEffect";

    case "폭포 차트":
      return "Waterfall";

    case "깔때기형 차트":
      return "Funnel";

    case "고가-저가-종가 차트":
      return "StockHLC";

    case "시가-고가-저가-종가 차트":
      return "StockOHLC";

    case "거래량-고가-저가-종가 차트":
      return "StockVHLC";

    case "거래량-시가-고가-저가-종가 차트":
      return "StockVOHLC";

    case "3차원 표면형 차트":
      return "Surface";

    case "3차원 표면형 차트(골격형)":
      return "SurfaceWireframe";

    case "표면형 차트(조감도)":
      return "SurfaceTopView";

    case "표면형 차트(골격형 조감도)":
      return "SurfaceTopViewWireframe";

    case "방사형 차트":
      return "Radar";

    case "표식이 있는 방사형 차트":
      return "RadarMarkers";

    case "채워진 방사형 차트":
      return "RadarFilled";

    case "원기둥 묶은 세로 막대형 차트":
      return "CylinderColClustered";

    case "원기둥 누적 세로 막대형 차트":
      return "CylinderColStacked";

    case "원기둥 100% 기준 누적 세로 막대형 차트":
      return "CylinderColStacked100";

    case "원기둥 묶은 가로 막대형 차트":
      return "CylinderBarClustered";

    case "원기둥 누적 가로 막대형 차트":
      return "CylinderBarStacked";

    case "원기둥 100% 기준 누적 가로 막대형 차트":
      return "CylinderBarStacked100";

    case "원기둥 세로 막대형 차트":
      return "CylinderCol";

    case "원뿔 묶은 세로 막대형 차트":
      return "ConeColClustered";

    case "원뿔 누적 세로 막대형 차트":
      return "ConeColStacked";

    case "원뿔 100% 기준 누적 세로 막대형 차트":
      return "ConeColStacked100";

    case "원뿔 묶은 가로 막대형 차트":
      return "ConeBarClustered";

    case "원뿔 누적 가로 막대형 차트":
      return "ConeBarStacked";

    case "원뿔 100% 기준 누적 가로 막대형 차트":
      return "ConeBarStacked100";

    case "원뿔 세로 막대형 차트":
      return "ConeCol";

    case "피라미드 묶은 세로 막대형 차트":
      return "PyramidColClustered";

    case "피라미드 누적 세로 막대형 차트":
      return "PyramidColStacked";

    case "피라미드 100% 기준 누적 세로 막대형 차트":
      return "PyramidColStacked100";

    case "피라미드 묶은 가로 막대형 차트":
      return "PyramidBarClustered";

    case "피라미드 누적 가로 막대형 차트":
      return "PyramidBarStacked";

    case "피라미드 100% 기준 누적 가로 막대형 차트":
      return "PyramidBarStacked100";

    case "피라미드 세로 막대형 차트":
      return "PyramidCol";

    case "지역 맵형 차트":
      return "RegionMap";

    default:
      return "Unknown chart type";
  }
}

async function evaluateTestFormula(newFormula: string) {
  try {
    let testResult = "";

    await Excel.run(async (context) => {
      const { workbook } = context;
      const originSheet = workbook.worksheets.getActiveWorksheet();

      try {
        workbook.worksheets.getItem("TestSheet").delete();
        await context.sync();
      } catch (error: unknown) {
        if (error instanceof Error) {
          throw error;
        }
      }
    });

    return testResult;
  } catch (e: unknown) {
    let warningMessage: MessageType = { type: "", title: "", body: "" };

    if (e instanceof Error) {
      warningMessage = {
        type: "warning",
        title: "에러 발생: ",
        body: `테스트를 진행 중 에러가 발생했습니다.${e.message}`,
      };
    }

    updateState("setMessageList", warningMessage);
  }
}

export {
  registerSelectionChange,
  getCellValue,
  updateState,
  splitCellAddress,
  extractAddresses,
  extractArgsAddress,
  getCellsInRange,
  nextColumn,
  activeSheetId,
  addPreset,
  deletePreset,
  getLastCellAddress,
  getTargetCellValue,
  getChartTypeInKorean,
  getChartTypeInEnglish,
  extractArgsFromFormula,
  extractFunctionsFromFormula,
  evaluateTestFormula,
};
