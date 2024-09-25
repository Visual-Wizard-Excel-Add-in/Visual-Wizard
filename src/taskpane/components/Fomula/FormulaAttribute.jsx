import { useCallback } from "react";
import { Switch } from "@fluentui/react-components";

import useStore from "../../utils/store";
import { highlightingCell } from "../../utils/cellStyleFunc";
import { groupCellsIntoRanges } from "../../utils/cellFormulaFunc";
import { extractAddresses } from "../../utils/commonFuncs";

function FormulaAttribute() {
  const isCellHighlighting = useStore((state) => state.isCellHighlighting);
  const setIsCellHighlighting = useStore(
    (state) => state.setIsCellHighlighting,
  );
  const cellFormula = useStore((state) => state.cellFormula);
  const cellArguments = useStore((state) => state.cellArguments);
  const cellAddress = useStore((state) => state.cellAddress);
  const cellValue = useStore((state) => state.cellValue);
  const cellFunctions = useStore((state) => state.cellFunctions);

  const handleHighlighting = useCallback(async () => {
    const newHighlightState = !isCellHighlighting;

    highlightingCell(
      newHighlightState,
      extractAddresses(cellFormula),
      cellAddress,
    );

    setIsCellHighlighting(newHighlightState);
  }, [isCellHighlighting, cellFormula, cellAddress]);

  const groupedCellArguments =
    groupCellsIntoRanges(cellArguments.map((arg) => arg.split("(")[0])) || [];

  const formattedCellArguments = groupedCellArguments
    .map((arg) => {
      const matchingArg = cellArguments.find((ca) => ca.startsWith(arg));

      if (matchingArg && !arg.includes(":")) {
        return `${arg}(${matchingArg.split("(")[1]}`;
      }

      return arg;
    })
    .join(", ");

  const resultCellAddress = cellAddress.split("!")[1];

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
              ? `${resultCellAddress}(${cellValue})`
              : null}
          </span>
        </p>
      </div>
    </div>
  );
}
export default FormulaAttribute;
