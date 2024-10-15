import { useEffect, useCallback, useRef } from "react";

import usePublicStore from "../store/publicStore";
import Header from "./Header";
import Formula from "./Fomula/Formula";
import Style from "./Style/Style";
import Macro from "./Macro/Macro";
import Validity from "./Validate/Validity";
import Share from "./Share/Share";
import { useStyles } from "../utils/style";
import { addOnSelectionChange, updateCellInfo } from "../utils/commonFuncs";
import CustomMessageBar from "./common/CustomMessageBar";

function App() {
  const globalSheetChangeHandler = useRef(null);
  const category = usePublicStore((state) => state.category);
  const sheetId = usePublicStore((state) => state.sheetId);
  const setSheetId = usePublicStore((state) => state.setSheetId);
  const messageList = usePublicStore((state) => state.messageList);
  const styles = useStyles();

  const categories = {
    Formula: <Formula />,
    Style: <Style />,
    Macro: <Macro />,
    Validity: <Validity />,
    Share: <Share />,
  };
  const CurrentCategory = categories[category] || null;

  const onWorksheetChanged = useCallback(
    async (event) => {
      const currentSheetId = event.worksheetId;

      if (currentSheetId !== sheetId) {
        setSheetId(currentSheetId);
      }
    },
    [sheetId, setSheetId],
  );

  useEffect(() => {
    const initializeSelectionEvent = async () => {
      await Excel.run(async (context) => {
        const { worksheets } = context.workbook;
        const sheet = worksheets.getActiveWorksheet();

        sheet.load("id");
        await context.sync();

        setSheetId(sheet.id);

        if (sheetId) {
          await addOnSelectionChange(sheetId, updateCellInfo);
        }
      });
    };

    initializeSelectionEvent();
  }, [sheetId]);

  useEffect(() => {
    const removeExistingHandler = async () => {
      if (globalSheetChangeHandler.current) {
        try {
          await Excel.run(
            globalSheetChangeHandler.current.context,
            async (context) => {
              globalSheetChangeHandler.current.remove();
              await context.sync();
            },
          );
        } catch (error) {
          throw new Error("핸들러 제거 중 오류:", error.message);
        }
      }
    };

    const initializeSheetChange = async () => {
      await Excel.run(async (context) => {
        const { worksheets } = context.workbook;

        globalSheetChangeHandler.current =
          worksheets.onActivated.add(onWorksheetChanged);
        await context.sync();
      });
    };

    initializeSheetChange();

    return async () => {
      await removeExistingHandler();
    };
  }, [onWorksheetChanged]);

  return (
    <div className={styles.root}>
      <Header />
      {messageList.length !== 0 && <CustomMessageBar />}
      {CurrentCategory || ""}
    </div>
  );
}

export default App;
