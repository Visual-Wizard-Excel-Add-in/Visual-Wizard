import { Switch } from "@fluentui/react-components";

import usePublicStore from "../../store/publicStore";
import { highlightingCell } from "../../utils/cellStyleFuncs";
import { groupCellsIntoRanges } from "../../utils/formulaFuncs";

function FormulaAttribute() {
  const isCellHighlighting = usePublicStore(
    (state) => state.isCellHighlighting,
  );
  const setIsCellHighlighting = usePublicStore(
    (state) => state.setIsCellHighlighting,
  );
  const cellArguments = usePublicStore((state) => state.cellArguments);
  const cellAddress = usePublicStore((state) => state.cellAddress);
  const cellValue = usePublicStore((state) => state.cellValue);
  const cellFunctions = usePublicStore((state) => state.cellFunctions);

  const handleHighlighting = async () => {
    const newHighlightState = !isCellHighlighting;

    highlightingCell(newHighlightState, cellAddress);

    setIsCellHighlighting(newHighlightState);
  };

  const groupedCellArguments =
    groupCellsIntoRanges(cellArguments.map((arg) => arg.split("(")[0])) || [];

  const formattedCellArguments = groupedCellArguments
    .map((groupArg) => {
      const matchingArg = cellArguments.find((cellArg) =>
        cellArg.startsWith(groupArg),
      );

      if (matchingArg && !groupArg.includes(":")) {
        const value = matchingArg.split("(")[1].split(")")[0];
        const valueWithComma = getValueWithComma(value);

        return `${groupArg}(${valueWithComma})`;
      }

      return groupArg;
    })
    .join(", ");

  const resultCellAddress = cellAddress.split("!")[1];
  const resultCellValue = getValueWithComma(cellValue);

  function getValueWithComma(value) {
    if (typeof +value !== "number") {
      return value;
    }

    const valueInStr = typeof value === "string" ? value : String(cellValue);
    let valueWithComma = null;

    let endIndex = valueInStr.length;
    const valueArr = [];

    for (let i = valueInStr.length - 1; i >= 0; i -= 1) {
      if (endIndex - i === 3) {
        valueArr.push(valueInStr.slice(i, endIndex));

        endIndex = i;
      } else if (i === 0) {
        valueArr.push(valueInStr.slice(i, endIndex));
      }
    }

    valueWithComma = valueArr.reverse().join(",");

    return valueWithComma;
  }

  return (
    <div>
      <Switch
        label="현재 셀 강조하기"
        onChange={handleHighlighting}
        disabled={cellFunctions.length === 0}
      />
      <div>
        <p className="font-bold">
          <img
            src="src/taskpane/assets/highlightArgumentCell.png"
            alt="highlightArgCells"
            className="inline"
          />
          &nbsp;인수:&nbsp;
          <span className="font-normal">{formattedCellArguments}</span>
        </p>
        <p className="mb-2 font-bold">
          <img
            src="src/taskpane/assets/highlightResultCell.png"
            alt="highlightResultCells"
            className="inline"
          />
          &nbsp;결과:&nbsp;
          <span className="font-normal">
            {cellFunctions.length !== 0
              ? `${resultCellAddress}(${resultCellValue})`
              : null}
          </span>
        </p>
      </div>
    </div>
  );
}
export default FormulaAttribute;
