import { Switch } from "@fluentui/react-components";

import useStore from "../../utils/store";
import { highlightingCell } from "../../utils/cellStyleFunc";
import { groupCellsIntoRanges } from "../../utils/cellFormulaFunc";

function FormulaAttribute() {
  const { cellArguments, cellAddress, cellValue, cellFunctions } = useStore();

  async function handleHighlighting() {
    useStore.setState((state) => {
      const newHighlightState = !state.isCellHighlighting;
      highlightingCell(
        newHighlightState,
        state.cellArguments,
        state.cellAddress,
      );
      return { isCellHighlighting: newHighlightState };
    });
  }

  const groupedCellArguments = groupCellsIntoRanges(
    cellArguments.map((arg) => arg.split("(")[0]),
  );

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
