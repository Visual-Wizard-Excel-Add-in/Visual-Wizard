import { Switch } from "@fluentui/react-components";

function FormulaAttribute() {
  function highlightCells() {}

  return (
    <div>
      <Switch label="강조하기" onChange={highlightCells} />
      <div>
        <p className="font-bold">
          <img
            src="assets/highlightArgumentCell.png"
            alt="highlightArgCells"
            className="inline"
          />
          &nbsp;인수
          <span className="font-normal">: A1()</span>
        </p>
        <p className="mb-2 font-bold">
          <img
            src="assets/highlightResultCell.png"
            alt="highlightResultCells"
            className="inline"
          />
          &nbsp;결과
          <span className="font-normal">: A1()</span>
        </p>
      </div>
    </div>
  );
}
export default FormulaAttribute;
