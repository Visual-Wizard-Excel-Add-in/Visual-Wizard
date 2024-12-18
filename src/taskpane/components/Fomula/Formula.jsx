import FeatureTab from "../common/FeatureTab";
import useTotalStore from "../../store/useTotalStore";
import FormulaInformation from "./FormulaInfomation";
import FormulaAttribute from "./FormulaAttribute";
import FormulaOrder from "./FormulaOrder";

function Fomula() {
  const setOpenTab = useTotalStore((state) => state.setOpenTab);
  const openTab = useTotalStore((state) => state.openTab);
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

export default Fomula;
