import { initialize, getApi, SheetsError } from "@equalto/sheets";
import { readFileSync } from "fs";

describe("Workbook - Cell operations", () => {
  beforeAll(async () => {
    await initialize();
  });

  test("can read formula on empty cell - returns null", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.get(0);
    expect(sheet.cell("A1").formula).toBe(null);
  });

  test("can read formula on cell with value - returns null", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.get(0);
    sheet.cell("A1").value = "Hello world";
    expect(sheet.cell("A1").formula).toBe(null);
  });

  test("can evaluate formulas in cells", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.get("Sheet1");

    sheet.cell("A1").value = 13;
    sheet.cell("A2").formula = "=A1*3";
    expect(workbook.cell("Sheet1!A2").value).toEqual(13 * 3);
    expect(workbook.cell("Sheet1!A2").formula).toEqual("=A1*3");
  });

  test("cannot assign formula by value", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.get("Sheet1");

    sheet.cell("A1").value = 13;
    sheet.cell("A2").value = "=A1*3";
    expect(workbook.cell("Sheet1!A2").formula).toEqual(null);
    expect(workbook.cell("Sheet1!A2").value).toEqual("=A1*3");
  });

  test("can read formatted value from cell", async () => {
    const { loadWorkbookFromMemory } = await getApi();
    let xlsxFile = readFileSync("./api/xlsx/formats.xlsx");
    let workbook = loadWorkbookFromMemory(xlsxFile, "en", "Europe/Berlin");

    expect(workbook.cell("Sheet1!B2").value).toEqual(0);
    expect(workbook.cell("Sheet1!B2").formattedValue).toEqual("0%");

    expect(workbook.cell("Sheet1!C2").value).toEqual(1);
    expect(workbook.cell("Sheet1!C2").formattedValue).toEqual("100%");

    expect(workbook.cell("Sheet1!D2").value).toEqual(2);
    expect(workbook.cell("Sheet1!D2").formattedValue).toEqual("200%");

    expect(workbook.cell("Sheet1!E2").value).toEqual(-0.12);
    expect(workbook.cell("Sheet1!E2").formattedValue).toEqual("-12%");
  });

  test("throws when values are read on cell from deleted sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.add();
    const cell = sheet.cell("A1");
    cell.value = 7;

    sheet.delete();

    const failCase = () => {
      cell.value;
    };

    expect(failCase).toThrow("Could not find sheet with sheetId=2");
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when values are set on cell from deleted sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.add();
    const cell = sheet.cell("A1");

    sheet.delete();

    const failCase = () => {
      cell.value = 8;
    };

    expect(failCase).toThrow("Could not find sheet with sheetId=2");
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when formula is read on cell from deleted sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.add();
    sheet.cell("A1").value = 7;
    const cell = sheet.cell("A2");
    cell.formula = "=A1+3";
    expect(cell.value).toEqual(10);

    sheet.delete();

    const failCase = () => {
      cell.formula;
    };

    expect(failCase).toThrow("Could not find sheet with sheetId=2");
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when formula is set on cell from deleted sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.add();
    sheet.cell("A1").value = 7;
    const cell = sheet.cell("A2");
    sheet.delete();

    const failCase = () => {
      cell.formula = "=A1+3";
    };

    expect(failCase).toThrow("Could not find sheet with sheetId=2");
    expect(failCase).toThrow(SheetsError);
  });
});
