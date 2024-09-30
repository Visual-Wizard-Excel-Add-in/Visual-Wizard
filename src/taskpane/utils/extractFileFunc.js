import { popUpMessage } from "./commonFuncs";

async function executeFunction(selectedOption) {
  try {
    Excel.run(async (context) => {
      const optionSheetName = "SelectExtract";
      const completeSheetName = "TriggerComplete";
      let optionSheet =
        context.workbook.worksheets.getItemOrNullObject(optionSheetName);
      let completeSheet =
        context.workbook.worksheets.getItemOrNullObject(completeSheetName);

      await context.sync();

      if (!optionSheet.isNullObject || !completeSheet.isNullObject) {
        optionSheet?.delete();
        completeSheet?.delete();
        await context.sync();
      }

      optionSheet = context.workbook.worksheets.add(optionSheetName);

      const targetRange = optionSheet.getRange("A1");

      targetRange.values = [[`${selectedOption}`]];

      completeSheet = context.workbook.worksheets.add(completeSheetName);
      await context.sync();

      setTimeout(async () => {
        optionSheet.delete();
        completeSheet.delete();
        await context.sync();
      }, 1000);
    });
  } catch (error) {
    popUpMessage("workFail", error.message);
  }
}
export default executeFunction;
