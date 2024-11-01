import { useState, useEffect } from "react";
import { Button, Input, Divider } from "@fluentui/react-components";

import useTotalStore from "../../store/useTotalStore";
import { evaluateTestFormula } from "../../utils/validateFuncs";

function FormulaTest() {
  const [args, setArgs] = useState([]);
  const [inputValues, setInputValues] = useState({});
  const [testResult, setTestResult] = useState(null);
  const cellFormula = useTotalStore((state) => state.cellFormula);
  const cellValue = useTotalStore((state) => state.cellValue);
  const cellArguments = useTotalStore((state) => state.cellArguments);
  const cellAddress = useTotalStore((state) => state.cellAddress);

  useEffect(() => {
    const fetchArgs = async () => {
      if (cellFormula) {
        const pureArgs = cellArguments.map((cellArg) => {
          if (cellArg.split("!")[0] === cellAddress.split("!")[0]) {
            return cellArg.split("!")[1];
          }
          return cellArg;
        });

        setArgs(pureArgs);
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

  const handleExecute = async () => {
    let newFormula = cellFormula;

    Object.entries(inputValues).forEach(([arg, value]) => {
      newFormula = newFormula.replace(arg, value);
    });

    const result = await evaluateTestFormula(newFormula);

    setTestResult(result);
  };

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
          {`${index + 1}. ${arg}`}
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
