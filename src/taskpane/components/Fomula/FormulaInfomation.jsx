import { Link } from "@fluentui/react-components";
import FORMULA_EXPLANATION from "../../utils/formulaExplanation";
import { useStyles } from "../../utils/style";

function FormulaInformation({ currentFormula }) {
  const styles = useStyles();

  return (
    <div>
      {currentFormula.map((formula, index) => {
        const func = Object.keys(formula)[0];
        return (
          <div key={{ formula } + { index }}>
            <p className="font-bold">{func}</p>
            <span>: {FORMULA_EXPLANATION[func]}</span>
            <hr className={styles.border} />
          </div>
        );
      })}
      <p className={styles.blurText}>
        자세한 설명은&nbsp;
        <Link
          appearance="inline"
          href="https://support.microsoft.com/ko-kr/office/excel-%ED%95%A8%EC%88%98-%EC%82%AC%EC%A0%84%EC%88%9C-b3944572-255d-4efb-bb96-c6d90033e188#bm19"
        >
          이곳
        </Link>
        을 참고해 주세요.
      </p>
    </div>
  );
}

export default FormulaInformation;
