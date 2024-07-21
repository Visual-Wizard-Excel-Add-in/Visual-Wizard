import { useEffect } from "react";

import useStore from "../utils/store";
import Header from "./Header";
import Formula from "./Fomula/Formula";
import Style from "./Style/Style";
import Macro from "./Macro/Macro";
import Validate from "./Validate/Validate";
import Share from "./Share/Share";
import { useStyles } from "../utils/style";
import {
  registerSelectionChange,
  getCellValue,
} from "../utils/cellCommonUtils";

function App() {
  const styles = useStyles();
  const { category, activeSheetName, sheetId, setSheetId } = useStore();

  const categories = {
    Formula: <Formula />,
    Style: <Style />,
    Macro: <Macro />,
    Validate: <Validate />,
    Share: <Share />,
  };
  const CurrentCategory = categories[category] || null;

  useEffect(() => {
    handleSheetChange();
  }, []);

  async function handleSheetChange() {
    await Excel.run(async (context) => {
      const { workbook } = context;

      workbook.worksheets.onActivated.add(async () => {
        await activeSheetName(sheetId);
        await registerSelectionChange(sheetId, getCellValue);
      });

      workbook.worksheets.onActivated.add(async (event) => {
        const newSheetId = event.worksheetId;
        if (newSheetId !== sheetId) {
          setSheetId(newSheetId);
          await registerSelectionChange(newSheetId, getCellValue);
        }
      });

      const initialSheet = workbook.worksheets.getActiveWorksheet();
      initialSheet.load("id");
      await context.sync();
      const initialSheetId = initialSheet.id;
      if (initialSheetId !== sheetId) {
        setSheetId(initialSheetId);
        await registerSelectionChange(initialSheetId, getCellValue);
      }
    });
  }

  return (
    <div className={styles.root}>
      <Header />
      {CurrentCategory || <div />}
    </div>
  );
}

export default App;
