import { useState, useEffect } from "react";
import { Switch } from "@fluentui/react-components";

import { getLastCellAddress } from "../../utils/validateFuncs";
import { detectErrorCell } from "../../utils/cellStyleFuncs";

function Validation() {
  const [isError, setIsError] = useState(true);
  const [lastCell, setLastCell] = useState("");

  useEffect(() => {
    let selectionChangeHandler = null;

    fetchLastCellAddress();

    Excel.run(async (context) => {
      selectionChangeHandler =
        context.workbook.worksheets.onSelectionChanged.add(
          fetchLastCellAddress,
        );
      await context.sync();
    });

    return removeFetchHandler;

    async function fetchLastCellAddress() {
      const address = await getLastCellAddress();

      setLastCell(address);
    }

    async function removeFetchHandler() {
      if (selectionChangeHandler) {
        await Excel.run(selectionChangeHandler.context, async (context) => {
          selectionChangeHandler.remove();
          await context.sync();
        });

        selectionChangeHandler = null;
      }
    }
  }, []);

  const highlightError = async () => {
    await detectErrorCell(isError);
    setIsError((prev) => !prev);
  };

  return (
    <div>
      <Switch label="에러 셀 검사" onChange={highlightError} />
      <p>
        사용중인 마지막 셀 영역:&nbsp;
        <span className="font-bold">{lastCell}</span>
      </p>
    </div>
  );
}

export default Validation;
