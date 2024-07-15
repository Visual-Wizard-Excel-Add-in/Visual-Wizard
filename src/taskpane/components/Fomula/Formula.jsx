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

  function handToggle(event, data) {
    setOpenTab(data.openItems);
  }

  return (
    <div className="mt-2">
      <FeatureTab
        order="1"
        featureName="정보"
        openTab={openTab}
        handToggle={handToggle}
        featureContents={FormulaInformation}
        currentFormula={currentFormula}
      />
      <FeatureTab
        order="2"
        featureName="참조"
        openTab={openTab}
        handToggle={handToggle}
        featureContents={FormulaAttribute}
        currentFormula={currentFormula}
      />
      <FeatureTab
        order="3"
        featureName="순서"
        openTab={openTab}
        handToggle={handToggle}
        featureContents={FormulaOrder}
        currentFormula={currentFormula}
      />
    </div>
  );
}

export default Fomula;
