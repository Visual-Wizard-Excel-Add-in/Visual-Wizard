import useStore from "../../utils/store";
import FeatureTab from "../common/FeatureTab";
import Mail from "./Mail";

function Share() {
  const { openTab, setOpenTab } = useStore();
  const features = [
    {
      name: "메일 전송",
      component: Mail,
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

export default Share;
