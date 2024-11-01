import useTotalStore from "../../store/useTotalStore";
import FeatureTab from "../common/FeatureTab";
import Extraction from "./Extraction";

function Share() {
  const [openTab, setOpenTab] = useTotalStore((state) => [
    state.openTab,
    state.setOpenTab,
  ]);
  const features = [
    {
      name: "추출하기",
      component: Extraction,
    },
  ];

  const handleToggle = (event, data) => {
    setOpenTab(data.openItems);
  };

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

export default Share;
