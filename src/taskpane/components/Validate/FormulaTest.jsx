import { Button, Field, Textarea } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";

function FormulaTest() {
  const selectedCellFormula = "SUM(A1, B1)";
  const styles = useStyles();
  const testResult = 15;

  return (
    <div>
      <p>
        선택한 셀의 수식:&nbsp;
        <span className="font-bold">{selectedCellFormula}</span>
      </p>
      <Field label="테스트 인수">
        <Textarea className="h-40" />
      </Field>
      <div className="flex justify-center mt-2">
        <Button size="small">실행</Button>
      </div>
      <hr className={styles.border} />
      <p className="text-xl font-bold">테스트 결과: {testResult}</p>
    </div>
  );
}

export default FormulaTest;
