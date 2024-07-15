import useStore from "../../utils/store";
import FeatureTab from "../common/FeatureTab";
import MacroRecord from "./MacroRecord";
import MacroSetting from "./MacroSetting";

function Macro() {
  const { openTab, setOpenTab } = useStore();
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

  function handToggle(event, data) {
    setOpenTab(data.openItems);
  }

  return (
    <div className="mt-2">
      {features.map((feature, index) => (
        <FeatureTab
          key={feature.name}
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

export default Macro;
