import { useCallback, useState } from "react";
import { v4 as uuidv4 } from "uuid";

import useStore from "../../utils/store";
import FeatureTab from "../common/FeatureTab";
import MacroRecord from "./MacroRecord";
import MacroSetting from "./MacroSetting";
import MacroNoticeBar from "./MacroNoticeBar";

function Macro() {
  const [openTab, setOpenTab] = useStore((state) => [
    state.openTab,
    state.setOpenTab,
  ]);
  const [isShowNoticeBar, setIsShowNoticeBar] = useState(true);
  const features = [
    {
      name: "매크로 녹화",
      component: MacroRecord,
    },
    {
      name: "매크로 설정",
      component: MacroSetting,
    },
  ];

  const handleToggle = useCallback((event, data) => {
    setOpenTab(data.openItems);
  }, []);

  return (
    <>
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
      <div className="flex justify-center">
        {isShowNoticeBar && (
          <MacroNoticeBar setIsShowNoticeBar={setIsShowNoticeBar} />
        )}
      </div>
    </>
  );
}

export default Macro;
