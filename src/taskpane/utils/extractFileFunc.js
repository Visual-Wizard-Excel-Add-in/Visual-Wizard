import { updateState } from "./cellCommonUtils";

async function executeFunction(selectedOption) {
  try {
    Excel.run(async (context) => {
      const sheetName = "SelctExtract1";
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);

      await context.sync();

      sheet.delete();
      await context.sync();

      sheet = context.workbook.worksheets.add(sheetName);
      await context.sync();

      sheet.getRange("A1").values = [[`${selectedOption}`]];
      await context.sync();

      sheet.delete();
      await context.sync();
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
