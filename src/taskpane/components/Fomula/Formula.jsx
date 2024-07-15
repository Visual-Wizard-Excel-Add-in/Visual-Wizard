import FeatureTab from "../common/FeatureTab";
import useStore from "../../utils/store";
import FormulaInformation from "./FormulaInfomation";
import FormulaAttribute from "./FormulaAttribute";
import FormulaOrder from "./FormulaOrder";

function Fomula() {
  const { openTab, setOpenTab } = useStore();
  const currentFormula = [
    {
      IF: `조건: AND( E5>D5, E5<Q$2)\ntrue: '회기오류'\nfalse: 2.IF 결과: 거짓`,
    },
    { AND: "" },
    { OR: "" },
    { DATEIF: "" },
  ];

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
          currentFormula={currentFormula}
        />
      ))}
    </div>
  );
}

export default Fomula;
