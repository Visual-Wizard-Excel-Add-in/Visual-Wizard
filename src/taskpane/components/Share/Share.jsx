import { useCallback } from "react";
import { v4 as uuidv4 } from "uuid";

import useStore from "../../utils/store";
import FeatureTab from "../common/FeatureTab";
import Extraction from "./Extraction";

function Share() {
  const [openTab, setOpenTab] = useStore((state) => [
    state.openTab,
    state.setOpenTab,
  ]);
  const features = [
    {
      name: "추출하기",
      component: Extraction,
    },
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

export default Share;
