import { initialize } from "@equalto/sheets";

async function run() {
  const { Workbook } = await initialize();

  let workbook = Workbook.loadFromFileSync("./xlsx/test.xlsx");

  workbook.evaluate();
  console.log(`B1=${workbook.getFormattedCellValue(0, 1, 2)}`);
  console.log(`B2=${workbook.getFormattedCellValue(0, 2, 2)}`);
  console.log(`B3=${workbook.getFormattedCellValue(0, 3, 2)}`);

  console.log("Change B1 to 6");
  workbook.setInput(0, 1, 2, "6");

  workbook.evaluate();
  console.log(`B1=${workbook.getFormattedCellValue(0, 1, 2)}`);
  console.log(`B2=${workbook.getFormattedCellValue(0, 2, 2)}`);
  console.log(`B3=${workbook.getFormattedCellValue(0, 3, 2)}`);

  // Exception demo
  console.log("");
  console.log("---");
  console.log("Exception demo");
  {
    try {
      // new Workbook("invalid_locale", "invalid_tz"); // Invalid timezone
      Workbook.loadFromMemory(new Uint8Array(0), "en", "Europe/Berlin"); // Invalid Zip Archive
    } catch (e) {
      handleWebAssemblyError(e);
    }
  }

  function handleWebAssemblyError(e) {
    console.log("error: ", typeof e);
    console.log("instanceof Error", e instanceof Error);
    if (e instanceof Error) {
      console.log("message: ", e.message);
      console.log("parsed message: ", JSON.parse(e.message));
      // next step: construct new error (wrap e for development purposes) and then reraise
    }
  }
}

run();
