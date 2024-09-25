import { updateState } from "./commonFuncs";

async function executeFunction(selectedOption) {
  try {
    Excel.run(async (context) => {
      const sheetName = "SelectExtract";
      const completeSheetName = "TriggerComplete";
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      let completeSheet =
        context.workbook.worksheets.getItemOrNullObject(completeSheetName);

      await context.sync();

      if (!sheet.isNullObject) {
        sheet.delete();
        await context.sync();
      }

      var a = 1;

      sheet = context.workbook.worksheets.add(sheetName);
      await context.sync();

      const targetRange = sheet.getRange("A1");
      await context.sync();

      targetRange.values = [[`${selectedOption}`]];
      await context.sync();

      if (!completeSheet.isNullObject) {
        sheet.delete();
        await context.sync();
      }

      completeSheet = context.workbook.worksheets.add(completeSheetName);
      await context.sync();

      setTimeout(async () => {
        sheet.delete();
        completeSheet.delete();
        await context.sync();
      }, 1000);
    });
  } catch (error) {
    updateState("setMessageList", {
      type: "warning",
      title: "추출 실패: ",
      body: `추출 과정에서 에러가 발생했습니다. ${error.message}`,
    });
  }
}
export default executeFunction;
