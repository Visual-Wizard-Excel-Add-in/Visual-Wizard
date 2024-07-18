import { Button } from "@fluentui/react-components";
import useStore from "../../utils/store";

import CustomPopover from "../common/CustomPopover";

function FormulaOrder() {
  function trigger(text) {
    return <Button>{text}</Button>;
  }
  const { cellFormulas } = useStore();

  return (
    <div>
      {cellFormulas.map((formula, index) => {
        const func = Object.keys(formula)[0];
        const description = formula[func];
        return (
          <div key={{ func } + { index }} className="mb-2">
            <span>{index + 1}. </span>
            <CustomPopover
              position="after"
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
