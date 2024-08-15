import { useState } from "react";
import { useStyles } from "../../utils/style";

function FormulaOrderDescription({ step }) {
  const [isExpandedCondition, setIsExpandedCondition] = useState(false);
  const [isExpandedTrueValue, setIsExpandedTrueValue] = useState(false);
  const [isExpandedFalseValue, setIsExpandedFalseValue] = useState(false);
  const [isExpandedFormula, setIsExpandedFormula] = useState(false);

  const styles = useStyles();

  function truncateText(text, maxLength, isExpanded) {
    if (!text || text.length <= maxLength || isExpanded) {
      return text;
    }

    return `${text.substring(0, maxLength)}...`;
  }

  const {
    functionName,
    condition: stepCondition,
    trueValue,
    falseValue,
    criteriaRange,
    criteria,
    formula,
  } = step;

  const renderCondition = (condition, maxLength, isExpanded, setIsExpanded) => (
    <div>
      <strong>조건: {truncateText(condition, maxLength, isExpanded)}</strong>
      {condition.length > maxLength && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsExpanded(!isExpanded)}
          >
            {isExpanded ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );

  const renderCriteriaRange = (range, maxLength, isExpanded, setIsExpanded) => (
    <div>
      <strong>조건 범위: {truncateText(range, maxLength, isExpanded)}</strong>
      {range.length > maxLength && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsExpanded(!isExpanded)}
          >
            {isExpanded ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );

  return (
    <div className="break-all">
      <strong>참조 셀: {step.address}</strong>
      <br />
      {[
        "IF",
        "IFS",
        "SUMIF",
        "COUNTIF",
        "AVERAGEIF",
        "IFERROR",
        "IFNA",
        "SWITCH",
        "CHOOSE",
      ].includes(functionName) && (
        <div>
          {stepCondition &&
            renderCondition(
              stepCondition,
              35,
              isExpandedCondition,
              setIsExpandedCondition,
            )}
          {trueValue && (
            <div>
              <strong>
                참일 때: {truncateText(trueValue, 35, isExpandedTrueValue)}
              </strong>
              {trueValue.length > 35 && (
                <div>
                  <button
                    className={styles.blurText}
                    onClick={() => setIsExpandedTrueValue(!isExpandedTrueValue)}
                  >
                    {isExpandedTrueValue ? "[간략히]" : "[더보기]"}
                  </button>
                </div>
              )}
            </div>
          )}
          {falseValue && (
            <div>
              <strong>
                거짓일 때: {truncateText(falseValue, 35, isExpandedFalseValue)}
              </strong>
              {falseValue.length > 35 && (
                <div>
                  <button
                    className={styles.blurText}
                    onClick={() =>
                      setIsExpandedFalseValue(!isExpandedFalseValue)
                    }
                  >
                    {isExpandedFalseValue ? "[간략히]" : "[더보기]"}
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      )}
      {[
        "SUMIF",
        "SUMIFS",
        "COUNTIF",
        "COUNTIFS",
        "AVERAGEIF",
        "AVERAGEIFS",
      ].includes(functionName) && (
        <div>
          {criteriaRange &&
            renderCriteriaRange(
              criteriaRange,
              35,
              isExpandedCondition,
              setIsExpandedCondition,
            )}
          {criteria &&
            renderCondition(
              criteria,
              35,
              isExpandedCondition,
              setIsExpandedCondition,
            )}
        </div>
      )}
      <br />
      <strong>식: {truncateText(formula, 35, isExpandedFormula)}</strong>
      {formula.length > 35 && (
        <div>
          <button
            className={styles.blurText}
            onClick={() => setIsExpandedFormula(!isExpandedFormula)}
          >
            {isExpandedFormula ? "[간략히]" : "[더보기]"}
          </button>
        </div>
      )}
    </div>
  );
}

export default FormulaOrderDescription;
