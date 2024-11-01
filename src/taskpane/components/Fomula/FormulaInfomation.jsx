import { Link, Divider } from "@fluentui/react-components";

import FORMULA_EXPLANATION from "../../constants/formulaConstants";
import useTotalStore from "../../store/useTotalStore";
import { useStyles } from "../../utils/style";

function FormulaInformation() {
  const cellFunctions = useTotalStore((state) => state.cellFunctions);
  const styles = useStyles();

  return (
    <div>
      {cellFunctions.length !== 0 &&
        cellFunctions.map((func) => {
          return (
            <div key={func}>
              <p className="font-bold">{func}</p>
              <span className="whitespace-pre-wrap">
                : {FORMULA_EXPLANATION[func]}
              </span>
              <Divider className="my-2" appearance="strong" />
            </div>
          );
        })}
      {cellFunctions.length !== 0 ? (
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
        "수식이 입력된 셀을 선택해주세요."
      )}
    </div>
  );
}

export default FormulaInformation;
