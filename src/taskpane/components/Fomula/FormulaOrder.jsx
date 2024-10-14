import { useEffect } from "react";

import usePublicStore from "../../store/publicStore";
import FormulaOrderDetail from "./FormulaOrderDetail";
import CustomPopover from "../common/CustomPopover";
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
