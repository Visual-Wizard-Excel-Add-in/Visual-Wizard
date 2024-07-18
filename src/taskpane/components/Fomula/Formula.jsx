import FeatureTab from "../common/FeatureTab";
import useStore from "../../utils/store";
import FormulaInformation from "./FormulaInfomation";
import FormulaAttribute from "./FormulaAttribute";
import FormulaOrder from "./FormulaOrder";

function Fomula() {
  const { openTab, setOpenTab } = useStore();
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

export default Fomula;
