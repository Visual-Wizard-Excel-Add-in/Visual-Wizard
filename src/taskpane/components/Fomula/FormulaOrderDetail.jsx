import { useState } from "react";

import { useStyles } from "../../utils/style";
import CONDITION_FUNCTIONS_LIST from "../../constants/conditionFuncsListConstans";

function truncateText(text, maxLength, isExpanded) {
  if (!text || text.length <= maxLength || isExpanded) {
    return text;
  }

  return `${text.substring(0, maxLength)}...`;
}

function FormulaOrderDetail({ step }) {
  const [isShowCondition, setIsShowCondition] = useState(false);
  const [isShowWhenTrue, setIsShowWhenTrue] = useState(false);
  const [isShowWhneFalse, setIsShowWhenFalse] = useState(false);
  const [isShowFormula, setIsShowFormula] = useState(false);
  const styles = useStyles();

  const showStates = {
    isShowCondition,
    setIsShowCondition,
    isShowWhenTrue,
    setIsShowWhenTrue,
    isShowWhneFalse,
    setIsShowWhenFalse,
    isShowFormula,
    setIsShowFormula,
  };

  const {
    functionName,
    condition,
    trueValue,
    falseValue,
    criteriaRange,
    criteria,
    formula,
  } = step;

  return (
    <div className="break-all">
      <p>
        <strong>참조 셀: </strong>
        {step.address}
      </p>
      {CONDITION_FUNCTIONS_LIST.single.includes(functionName) && (
        <SingleConditionDetail
          condition={condition}
          trueValue={trueValue}
          falseValue={falseValue}
          styles={styles}
          showStates={showStates}
        />
      )}
      {CONDITION_FUNCTIONS_LIST.serveral.includes(functionName) && (
        <ServeralConditionDetail
          criteriaRange={criteriaRange}
          criteria={criteria}
          styles={styles}
          showStates={showStates}
        />
      )}
      <FunctionFormula
        formula={formula}
        styles={styles}
        isShowFormula={isShowFormula}
        setIsShowFormula={setIsShowFormula}
      />
    </div>
  );
}

export default FormulaOrderDetail;

function SingleConditionDetail({
  condition,
  trueValue,
  falseValue,
  styles,
  showStates,
}) {
  return (
    <div>
      {condition && (
        <StepCondition
          condition={condition}
          maxLength={35}
          styles={styles}
          isShowCondition={showStates.isShowCondition}
          setIsShowCondition={showStates.setIsShowCondition}
        />
      )}
      {trueValue && (
        <WhenTrue
          trueValue={trueValue}
          styles={styles}
          isShowWhenTrue={showStates.isShowWhenTrue}
          setIsShowWhenTrue={showStates.setIsShowWhenTrue}
        />
      )}
      {falseValue && (
        <WhenFalse
          falseValue={falseValue}
          styles={styles}
          isShowWhneFalse={showStates.isShowWhneFalse}
          setIsShowWhenFalse={showStates.setIsShowWhenFalse}
        />
      )}
    </div>
  );
}

function ServeralConditionDetail({
  criteriaRange,
  criteria,
  styles,
  showStates,
}) {
  return (
    <div>
      {criteriaRange && (
        <StepConditionRange
          range={criteriaRange}
          maxLength={35}
          styles={styles}
          isShowCondition={showStates.isShowCondition}
          setIsShowCondition={showStates.setIsShowCondition}
        />
      )}
      {criteria && (
        <StepCondition
          condition={criteria}
          maxLength={35}
          styles={styles}
          isShowCondition={showStates.isShowCondition}
          setIsShowCondition={showStates.setIsShowCondition}
        />
      )}
    </div>
  );
}

function FunctionFormula({ formula, styles, isShowFormula, setIsShowFormula }) {
  return (
    <>
      <br />
      <p>
        <strong>식:</strong>
        {truncateText(formula, 35, isShowFormula)}
      </p>
      {formula.length > 35 && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsShowFormula(!isShowFormula)}
            type="button"
          >
            {isShowFormula ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </>
  );
}

function StepConditionRange({
  range,
  maxLength,
  styles,
  isShowCondition,
  setIsShowCondition,
}) {
  return (
    <div>
      <br />
      <p>
        <strong>조건 범위: </strong>
        {truncateText(range, maxLength, isShowCondition)}
      </p>
      {range.length > maxLength && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsShowCondition(!isShowCondition)}
            type="button"
          >
            {isShowCondition ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );
}

function StepCondition({
  condition,
  maxLength,
  styles,
  isShowCondition,
  setIsShowCondition,
}) {
  return (
    <div>
      <br />
      <p>
        <strong>조건: </strong>
        {truncateText(condition, maxLength, isShowCondition)}
      </p>
      {condition.length > maxLength && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsShowCondition(!isShowCondition)}
            type="button"
          >
            {isShowCondition ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );
}

function WhenTrue({ trueValue, styles, isShowWhenTrue, setIsShowWhenTrue }) {
  return (
    <div>
      <br />
      <strong>참일 때: {truncateText(trueValue, 35, isShowWhenTrue)}</strong>
      {trueValue.length > 35 && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsShowWhenTrue(!isShowWhenTrue)}
            type="button"
          >
            {isShowWhenTrue ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );
}

function WhenFalse({
  falseValue,
  styles,
  isShowWhneFalse,
  setIsShowWhenFalse,
}) {
  return (
    <div>
      <br />
      <p>
        <strong>거짓일 때:</strong>
        {truncateText(falseValue, 35, isShowWhneFalse)}
      </p>
      {falseValue.length > 35 && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsShowWhenFalse(!isShowWhneFalse)}
            type="button"
          >
            {isShowWhneFalse ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );
}
