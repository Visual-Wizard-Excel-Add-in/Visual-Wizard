import { useEffect } from "react";
import { Button } from "@fluentui/react-components";

import usePublicStore from "../../store/publicStore";
import CustomPopover from "../common/CustomPopover";
import FormulaOrderDescription from "./FormulaOrderDescription";
import { parseFormulaSteps } from "../../utils/formulaFuncs";
import { popUpMessage } from "../../utils/commonFuncs";

function FormulaOrder() {
  const cellFormula = usePublicStore((state) => state.cellFormula);
  const formulaSteps = usePublicStore((state) => state.formulaSteps);
  const setFormulaSteps = usePublicStore((state) => state.setFormulaSteps);

  useEffect(() => {
    async function fetchFormulaSteps() {
      try {
        if (cellFormula) {
          const result = await parseFormulaSteps(cellFormula);

          setFormulaSteps(result);
        } else {
          setFormulaSteps([]);
        }
      } catch (error) {
        setFormulaSteps([]);
        popUpMessage("load Failed", error.message);
      }
    }

    fetchFormulaSteps();
  }, [cellFormula]);

  return (
    <div>
      {formulaSteps && formulaSteps.length !== 0 ? (
        FormulaSteps(formulaSteps)
      ) : (
        <div>수식이 입력된 셀을 선택해주세요.</div>
      )}
    </div>
  );
}

function Trigger(Contents) {
  return <Button>{Contents}</Button>;
}

function FormulaSteps(steps) {
  return steps.map((step, index) => {
    const { functionName } = step;
    const description = <FormulaOrderDescription step={step} />;

    return (
      <div key={`${step.address}-${step.functionName}`}>
        <span>{index + 1}. </span>
        <CustomPopover
          position="after"
          triggerContents={Trigger(functionName)}
          PopoverContents={description}
        />
      </div>
    );
  });
}
export default FormulaOrder;
