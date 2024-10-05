import { useState, useEffect, useCallback } from "react";
import { Button, Input, Divider } from "@fluentui/react-components";

import usePublicStore from "../../store/publicStore";
import { extractReferenceCells } from "../../utils/commonFuncs";
import { evaluateTestFormula } from "../../utils/validateFuncs";
import {
  groupCellsIntoRanges,
  parseFormulaSteps,
} from "../../utils/formulaFuncs";

function FormulaTest() {
  const [args, setArgs] = useState([]);
  const [inputValues, setInputValues] = useState({});
  const [testResult, setTestResult] = useState(null);
  const cellFormula = usePublicStore((state) => state.cellFormula);
  const cellValue = usePublicStore((state) => state.cellValue);
  const cellArguments = usePublicStore((state) => state.cellArguments);

  useEffect(() => {
    const fetchArgs = async () => {
      if (cellFormula) {
        const formulaSteps = await parseFormulaSteps(cellFormula);
        const allArgs = formulaSteps.flatMap((step) => {
          const addresses = extractReferenceCells(step.address);

          return groupCellsIntoRanges(addresses);
        });
        const uniqueArgs = [...new Set(allArgs)];

        setArgs(uniqueArgs);
      } else {
        setArgs([]);
      }
    };

    fetchArgs();
    setTestResult(null);
  }, [cellFormula]);

  function handleInputChange(arg, value) {
    setInputValues((prevState) => ({ ...prevState, [arg]: value }));
  }

  const handleExecute = useCallback(async () => {
    let newFormula = cellFormula;

    Object.entries(inputValues).forEach(([arg, value]) => {
      newFormula = newFormula.replace(arg, value);
    });

    const result = await evaluateTestFormula(newFormula);

    setTestResult(result);
  }, [cellFormula, inputValues]);

  return (
    <div>
      <div>
        <p>선택한 셀의 수식: </p>
        <span className="inline font-bold break-words whitespace-pre-wrap">
          {cellFormula}
        </span>
        <p className="mt-2">
          현재 결과:&nbsp;
          <span className="font-bold">{cellValue}</span>
        </p>
      </div>
      <Divider className="my-2" appearance="strong" />
      {args.map((arg, index) => (
        <p key={arg} className="mb-2">
          {index + 1}. 인자:
          {cellArguments?.find((detailArg) => detailArg.includes(arg)) || arg}
          <br />
          <Input
            className="mt-1"
            onChange={(e) => handleInputChange(arg, e.target.value)}
            placeholder="변경할 값이나 셀 주소"
          />
        </p>
      ))}
      <div className="flex justify-center mt-2">
        <Button onClick={handleExecute} size="small">
          실행
        </Button>
      </div>
      <Divider className="my-2" appearance="strong" />
      <p className="text-xl font-bold">테스트 결과: {testResult}</p>
    </div>
  );
}

export default FormulaTest;
