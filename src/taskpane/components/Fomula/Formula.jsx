import { useCallback } from "react";
import { v4 as uuidv4 } from "uuid";

import FeatureTab from "../common/FeatureTab";
import useStore from "../../utils/store";
import FormulaInformation from "./FormulaInfomation";
import FormulaAttribute from "./FormulaAttribute";
import FormulaOrder from "./FormulaOrder";

function Fomula() {
  const openTab = useStore((state) => state.openTab);
  const setOpenTab = useStore((state) => state.setOpenTab);
  const features = [
    {
      name: "정보",
      component: FormulaInformation,
    },
    {
      name: "참조",
      component: FormulaAttribute,
    },
    { name: "순서", component: FormulaOrder },
  ];

  const handleToggle = useCallback((event, data) => {
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
          handleToggle={handleToggle}
          featureContents={feature.component}
        />
      ))}
    </div>
  );
}

export default Fomula;
