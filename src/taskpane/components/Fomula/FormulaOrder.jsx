import { useEffect } from "react";
import { Button } from "@fluentui/react-components";
import { v4 as uuidv4 } from "uuid";

import useStore from "../../utils/store";
import CustomPopover from "../common/CustomPopover";
import FormulaOrderDescription from "./FormulaOrderDescription";
import { parseFormulaSteps } from "../../utils/cellFormulaFunc";

function FormulaOrder() {
  const cellFormula = useStore((state) => state.cellFormula);
  const formulaSteps = useStore((state) => state.formulaSteps);
  const setFormulaSteps = useStore((state) => state.setFormulaSteps);

  useEffect(() => {
    async function fetchFormulaSteps() {
      if (cellFormula) {
        try {
          const result = await parseFormulaSteps();

          setFormulaSteps(result);
        } catch (error) {
          setFormulaSteps([]);
        }
      } else {
        setFormulaSteps([]);
      }
    }

    fetchFormulaSteps();
  }, [cellFormula]);

  function trigger(text) {
    return <Button>{text}</Button>;
  }

  function renderFormulaSteps(steps) {
    return steps.map((step, index) => {
      const func = step.functionName;
      const description = <FormulaOrderDescription step={step} />;
      return (
        <div key={uuidv4()}>
          <span>{index + 1}. </span>
          <CustomPopover
            position="after"
            triggerContents={trigger(func)}
            PopoverContents={description}
          />
        </div>
      );
    });
  }

  return (
    <div>
      {formulaSteps && formulaSteps.length !== 0 ? (
        renderFormulaSteps(formulaSteps)
      ) : (
        <div>수식이 입력된 셀을 선택해주세요.</div>
      )}
    </div>
  );
}

export default FormulaOrder;
