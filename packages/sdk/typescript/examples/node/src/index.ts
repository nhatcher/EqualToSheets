import { readFileSync } from "fs";
import { initialize, IWorkbook } from "@equalto/sheets";

function calculateForInput(workbook: IWorkbook, inputValue: number) {
  workbook.cell("Sheet1!B1").value = inputValue;

  console.log(`Results for ${inputValue}`);
  console.log(
    workbook.cell("Sheet1!A2").value,
    workbook.cell("Sheet1!B2").formattedValue
  );
  console.log(
    workbook.cell("Sheet1!A3").value,
    workbook.cell("Sheet1!B3").formattedValue
  );
  console.log(
    workbook.cell("Sheet1!A4").value,
    workbook.cell("Sheet1!B4").formattedValue
  );
}

async function run() {
  const { loadWorkbookFromMemory } = await initialize();

  let xlsxFile = readFileSync("./xlsx/test.xlsx");
  let workbook = loadWorkbookFromMemory(xlsxFile, "en", "Europe/Berlin");

  calculateForInput(workbook, 3);
  calculateForInput(workbook, 7);
}

run();
