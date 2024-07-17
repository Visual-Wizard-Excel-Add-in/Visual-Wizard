import { Link } from "@fluentui/react-components";

import FORMULA_EXPLANATION from "../../utils/formulaExplanation";
import useStore from "../../utils/store";
import { useStyles } from "../../utils/style";

function FormulaInformation() {
  const styles = useStyles();
  const { cellFormulas } = useStore();

  return (
    <div>
      {cellFormulas.map((formula) => {
        return (
          <div key={formula}>
            <p className="font-bold">{formula}</p>
            <span className="whitespace-pre-wrap">
              : {FORMULA_EXPLANATION[formula]}
            </span>
            <hr className={styles.border} />
          </div>
        );
      })}
      {cellFormulas.length !== 0 ? (
        <p className={styles.blurText}>
          자세한 설명은&nbsp;
          <Link
            className={styles.fontBolder}
            appearance="inline"
            href="https://support.microsoft.com/ko-kr/office/excel-%ED%95%A8%EC%88%98-%EC%82%AC%EC%A0%84%EC%88%9C-b3944572-255d-4efb-bb96-c6d90033e188#bm19"
          >
            이곳
          </Link>
          을 참고해 주세요.
        </p>
      ) : (
        "수식이 아닙니다."
      )}
    </div>
  );
}

export default FormulaInformation;
