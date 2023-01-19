import { initialize, getApi } from "@equalto/sheets";

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
});
