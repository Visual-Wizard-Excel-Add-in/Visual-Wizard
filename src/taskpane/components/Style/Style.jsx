import { useCallback } from "react";

import usePublicStore from "../../store/publicStore";
import FeatureTab from "../common/FeatureTab";
import CellStyle from "./CellStyle";
import ChartStyle from "./ChartStyle";

function Style() {
  const [openTab, setOpenTab] = usePublicStore((state) => [
    state.openTab,
    state.setOpenTab,
  ]);
  const features = [
    {
      name: "셀 서식",
      component: CellStyle,
    },
    {
      name: "차트 서식",
      component: ChartStyle,
    },
  ];

  const handleToggle = useCallback((event, data) => {
    setOpenTab(data.openItems);
  }, []);

  return (
    <div className="mt-2">
      {features.map((feature, index) => (
        <FeatureTab
          key={feature.name}
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

export default Style;
