import { readFileSync } from "fs";
import { initialize, IWorkbook } from "@equalto/sheets";

const printSheets = (workbook: IWorkbook) => {
  console.log(
    "Existing sheets: ",
    workbook.sheets.all().map((sheet) => ({
      index: sheet.index,
      id: sheet.id,
      name: sheet.name,
    }))
  );
};

async function run() {
  const { loadWorkbookFromMemory } = await initialize();

  let xlsxFile = readFileSync("./xlsx/test.xlsx");
  let workbook = loadWorkbookFromMemory(xlsxFile, "en", "Europe/Berlin");
  console.log("Workbook loaded.");

  let namedNewSheet = workbook.sheets.add("newSheet");
  console.log("Added sheet: ", {
    name: namedNewSheet.name,
    id: namedNewSheet.id,
    index: namedNewSheet.index,
  });

  let defaultNewSheet = workbook.sheets.add();
  console.log("Added sheet:", {
    name: defaultNewSheet.name,
    id: defaultNewSheet.id,
    index: defaultNewSheet.index,
  });

  printSheets(workbook);

  workbook.sheets.get(2).delete();
  console.log("Removed sheet with index=2");
  printSheets(workbook);

  let newSheetName: string = "renamedSheet";
  console.log(`Changing name '${defaultNewSheet.name}' to '${newSheetName}'`);
  defaultNewSheet.name = newSheetName;
  printSheets(workbook);

  let calcSheet = workbook.sheets.get(0);

  let cells = [
    calcSheet.cell("A1"), // A1
    calcSheet.cell("B1"), // B1
    calcSheet.cell(2, 2), // B2
    calcSheet.cell("B3"), // B3
    calcSheet.cell(4, 2), // B4
    calcSheet.cell(5, 2), // B5
  ];

  console.log(
    cells.map((cell) => `(${cell.row},${cell.column}) = ${cell.value}`)
  );

  console.log("Change B1 to 6");
  calcSheet.cell("B1").value = 6;

  console.log(
    cells.map((cell) => `(${cell.row},${cell.column}) = ${cell.value}`)
  );

  console.log(
    `B1 should not have formula, formula: ${workbook.cell("Sheet1!B1").formula}`
  );

  console.log(
    `Change B2 formula to =B1*B1*B1, (old: ${
      workbook.cell("Sheet1!B2").formula
    })`
  );
  calcSheet.cell("B2").formula = "=B1*B1*B1";

  console.log(
    cells.map((cell) => `(${cell.row},${cell.column}) = ${cell.value}`)
  );

  // Exception demo
  console.log("");
  console.log("---");
  console.log("Exception demo");
  {
    try {
      // new Workbook("invalid_locale", "invalid_tz"); // Invalid timezone
      loadWorkbookFromMemory(new Uint8Array(0), "en", "Europe/Berlin"); // Invalid Zip Archive
    } catch (e) {
      handleWebAssemblyError(e);
    }
  }

  function handleWebAssemblyError(e: any) {
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
