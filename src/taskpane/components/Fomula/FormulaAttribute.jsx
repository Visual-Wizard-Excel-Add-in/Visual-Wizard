import { Switch } from "@fluentui/react-components";
import { useCallback, useEffect, useState } from "react";

import usePublicStore from "../../store/publicStore";
import { targetCellValue } from "../../utils/commonFuncs";
import { highlightingCell } from "../../utils/cellStyleFuncs";

function FormulaAttribute() {
  const cellArguments = usePublicStore((state) => state.cellArguments);
  const cellAddress = usePublicStore((state) => state.cellAddress);
  const cellValue = usePublicStore((state) => state.cellValue);
  const cellFunctions = usePublicStore((state) => state.cellFunctions);
  const isHighlight = usePublicStore((state) => state.isHighlight);
  const setIsHighlight = usePublicStore((state) => state.setIsHighlight);
  const [argsWithValue, setArgsWithValue] = useState("");

  const convertUnit = useCallback((value) => {
    return !+value ? value : new Intl.NumberFormat("ko-KR").format(value);
  }, []);

  useEffect(() => {
    const fetchArgsWithValue = async () => {
      const results = await Promise.all(
        cellArguments.map(
          async (referCell) => await makeCellWithValue(referCell),
        ),
      );

      setArgsWithValue(results.join(", "));
    };

    async function makeCellWithValue(referCell) {
      const address =
        referCell.split("!")[0] === cellAddress.split("!")[0]
          ? referCell.split("!")[1]
          : referCell;

      if (!referCell.includes(":")) {
        const valueWithComma = convertUnit(await targetCellValue(referCell));

        return `${address}(${valueWithComma})`;
      }

      return address;
    }

    fetchArgsWithValue();
  }, [cellArguments, convertUnit]);

  const handleHighlighting = async () => {
    await highlightingCell(!isHighlight, cellAddress);

    setIsHighlight();
  };

  const resultCellWithValue = argsWithValue
    ? `${cellAddress.split("!")[1]}(${convertUnit(cellValue)})`
    : "";

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
          <span className="font-normal">{argsWithValue}</span>
        </p>
        <p className="mb-2 font-bold">
          <img
            src="src/taskpane/assets/highlightResultCell.png"
            alt="highlightResultCells"
            className="inline"
          />
          &nbsp;결과:&nbsp;
          <span className="font-normal">{resultCellWithValue}</span>
        </p>
      </div>
    </div>
  );
}
export default FormulaAttribute;
