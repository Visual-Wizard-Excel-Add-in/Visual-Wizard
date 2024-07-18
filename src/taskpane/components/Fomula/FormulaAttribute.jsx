import { Switch } from "@fluentui/react-components";

import useStore from "../../utils/store";
import { highlightingCell } from "../../utils/funcUtils";

function FormulaAttribute() {
  const { cellArguments, cellAddress, cellValue, cellFormulas } = useStore();

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

  return (
    <div>
      <Switch
        label="현재 셀 강조하기"
        onChange={handleHighlighting}
        disabled={cellFormulas.length === 0}
      />
      <div>
        <p className="font-bold">
          <img
            src="assets/highlightArgumentCell.png"
            alt="highlightArgCells"
            className="inline"
          />
          &nbsp;인수:&nbsp;
          <span className="font-normal">{cellArguments.join(", ")}</span>
        </p>
        <p className="mb-2 font-bold">
          <img
            src="assets/highlightResultCell.png"
            alt="highlightResultCells"
            className="inline"
          />
          &nbsp;결과:&nbsp;
          <span className="font-normal">
            {cellFormulas.length !== 0 ? `${cellAddress}(${cellValue})` : null}
          </span>
        </p>
      </div>
    </div>
  );
}
export default FormulaAttribute;
