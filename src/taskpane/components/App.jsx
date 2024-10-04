import { useEffect } from "react";

import useStore from "../utils/store";
import Header from "./Header";
import Formula from "./Fomula/Formula";
import Style from "./Style/Style";
import Macro from "./Macro/Macro";
import Validate from "./Validate/Validate";
import Share from "./Share/Share";
import { useStyles } from "../utils/style";
import { registerSelectionChange, updateCellInfo } from "../utils/commonFuncs";
import CustomMessageBar from "./common/CustomMessageBar";

let handleSheetChange = null;

function App() {
  const category = useStore((state) => state.category);
  const sheetId = useStore((state) => state.sheetId);
  const setSheetId = useStore((state) => state.setSheetId);
  const messageList = useStore((state) => state.messageList);
  const styles = useStyles();

  const categories = {
    Formula: <Formula />,
    Style: <Style />,
    Macro: <Macro />,
    Validate: <Validate />,
    Share: <Share />,
  };
  const CurrentCategory = categories[category] || null;

  useEffect(() => {
    Excel.run(async (context) => {
      const { worksheets } = context.workbook;
      const sheet = worksheets.getActiveWorksheet();

      sheet.load("id");
      await context.sync();

      setSheetId(sheet.id);

      registerSelectionChange(sheet.id, updateCellInfo);

      handleSheetChange = worksheets.onActivated.add((event) =>
        onWorksheetChanged(event),
      );

      await context.sync();
    });

    return () => {
      if (handleSheetChange !== null) {
        Excel.run(handleSheetChange.context, async (context) => {
          handleSheetChange?.remove();
          await context.sync();
        });
      }

      handleSheetChange = null;
    };
  }, []);

  function onWorksheetChanged(event) {
    const currentSheetId = event.worksheetId;

    if (currentSheetId !== sheetId) {
      setSheetId(currentSheetId);

      registerSelectionChange(currentSheetId, updateCellInfo);
    }
  }

  return (
    <div className={styles.root}>
      <Header />
      {messageList.length !== 0 && <CustomMessageBar />}
      {CurrentCategory || ""}
    </div>
  );
}

export default App;
