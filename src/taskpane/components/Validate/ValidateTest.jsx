import { useState, useEffect } from "react";
import { Switch } from "@fluentui/react-components";

import { getLastCellAddress } from "../../utils/cellCommonUtils";
import { detectErrorCell } from "../../utils/cellStyleFunc";

function ValidateTest() {
  const [isError, setIsError] = useState(true);
  const [lastCell, setLastCell] = useState("");

  useEffect(() => {
    const fetchLastCellAddress = async () => {
      const address = await getLastCellAddress();

      setLastCell(address);
    };

    fetchLastCellAddress();
  }, [lastCell]);

  async function highlightError() {
    await detectErrorCell(isError);
    setIsError((prev) => !prev);
  }

  return (
    <div>
      <Switch
        label="에러 셀 검사"
        onChange={() => {
          highlightError();
        }}
      />
      <p>
        사용중인 마지막 셀 영역:&nbsp;
        <span className="font-bold">{lastCell}</span>
      </p>
    </div>
  );
}

export default ValidateTest;
