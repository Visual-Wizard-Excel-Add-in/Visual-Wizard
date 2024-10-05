import usePublicStore from "../../store/publicStore";
import FeatureTab from "../common/FeatureTab";
import ValidateTest from "./ValidateTest";
import FormulaTest from "./FormulaTest";

function Validate() {
  const [openTab, setOpenTab] = usePublicStore((state) => [
    state.openTab,
    state.setOpenTab,
  ]);
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

export default Validate;
