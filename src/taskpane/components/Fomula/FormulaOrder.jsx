import { useEffect } from "react";

import useTotalStore from "../../store/useTotalStore";
import FormulaOrderDetail from "./FormulaOrderDetail";
import CustomPopover from "../common/CustomPopover";
import { parseFormulaSteps } from "../../utils/formulaFuncs";
import { popUpMessage } from "../../utils/commonFuncs";

function FormulaOrder() {
  const cellFormula = useTotalStore((state) => state.cellFormula);
  const formulaSteps = useTotalStore((state) => state.formulaSteps);
  const setFormulaSteps = useTotalStore((state) => state.setFormulaSteps);

  useEffect(() => {
    try {
      fetchFormulaSteps();
    } catch (error) {
      setFormulaSteps([]);
      popUpMessage("load Failed", error.message);
    }

    async function fetchFormulaSteps() {
      if (cellFormula) {
        const result = await parseFormulaSteps(cellFormula);

        setFormulaSteps(result);
      } else {
        setFormulaSteps([]);
      }
    }
  }, [cellFormula]);

  return (
    <div>
      {formulaSteps && formulaSteps.length !== 0 ? (
        <FormulaSteps steps={formulaSteps} />
      ) : (
        <div>수식이 입력된 셀을 선택해주세요.</div>
      )}
    </div>
  );
}

export default FormulaOrder;

function FormulaSteps({ steps }) {
  return (
    <>
      {steps.map((step, index) => {
        return (
          <div key={`${step.address}-${step.functionName}`}>
            <span>{index + 1}. </span>
            <CustomPopover
              position="after"
              triggerContents={step.functionName}
              PopoverContents={<FormulaOrderDetail step={step} />}
            />
          </div>
        );
      })}
    </>
  );
}
