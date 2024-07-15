import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import CustomPopover from "../common/CustomPopover";

function FormulaOrder({ currentFormula }) {
  const styles = useStyles();
  function trigger(text) {
    return <Button>{text}</Button>;
  }

  return (
    <div>
      {currentFormula.map((formula, index) => {
        const func = Object.keys(formula)[0];
        const description = formula[func];
        return (
          <div key={{ func } + { index }} className="mb-2">
            <span>{index + 1}. </span>
            <CustomPopover
              position="below"
              triggerContents={trigger(func)}
              PopoverContents={description}
            />
          </div>
        );
      })}
    </div>
  );
}

export default FormulaOrder;
