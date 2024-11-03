import { popUpMessage } from "./commonFuncs";

async function executeFunction(selectedOption) {
  const optionSheetName = "SelectExtract";
  const completeSheetName = "TriggerComplete";

  try {
    Excel.run(async (context) => {
      await checkRemainTrigger(context);

      const optionSheet = context.workbook.worksheets.add(optionSheetName);
      const targetRange = optionSheet.getRange("A1");

      targetRange.values = [[`${selectedOption}`]];

      const completeSheet = context.workbook.worksheets.add(completeSheetName);

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

  async function checkRemainTrigger(context) {
    const optionSheet =
      context.workbook.worksheets.getItemOrNullObject(optionSheetName);
    const completeSheet =
      context.workbook.worksheets.getItemOrNullObject(completeSheetName);

    await context.sync();

    if (!optionSheet.isNullObject) {
      optionSheet.delete();
    }

    if (!completeSheet.isNullObject) {
      completeSheet.delete();
    }

    await context.sync();
  }
}
export default executeFunction;
