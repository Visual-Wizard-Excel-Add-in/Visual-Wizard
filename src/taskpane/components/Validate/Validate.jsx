import { useCallback } from "react";
import { v4 as uuidv4 } from "uuid";

import useStore from "../../utils/store";
import FeatureTab from "../common/FeatureTab";
import ValidateTest from "./ValidateTest";
import FormulaTest from "./FormulaTest";

function Validate() {
  const openTab = useStore((state) => state.openTab);
  const setOpenTab = useStore((state) => state.setOpenTab);
  const features = [
    {
      name: "유효성 검사",
      component: ValidateTest,
    },
    {
      name: "수식 테스트",
      component: FormulaTest,
    },
  ];

  const handToggle = useCallback((event, data) => {
    setOpenTab(data.openItems);
  }, []);

  return (
    <div className="mt-2">
      {features.map((feature, index) => (
        <FeatureTab
          key={uuidv4()}
          order={String(index + 1)}
          featureName={feature.name}
          openTab={openTab}
          handToggle={handToggle}
          featureContents={feature.component}
        />
      ))}
    </div>
  );
}

export default Validate;
