import { initialize, getApi } from "@equalto/sheets";
import type { ISheet } from "@equalto/sheets";

const mapSheetToObject = (sheet: ISheet) => {
  return {
    id: sheet.id,
    index: sheet.index,
    name: sheet.name,
  };
};

describe("Workbook", () => {
  beforeAll(async () => {
    await initialize();
  });

  test("can list sheets in new workbook", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
    ]);
  });

  test("can create sheet with default name in new workbook", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const newSheet = workbook.sheets.add();

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
      { id: 2, index: 1, name: "Sheet2" },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual("Sheet2");
  });

  test("can create sheet with given name in new workbook", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const newSheet = workbook.sheets.add("MyName");

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
      { id: 2, index: 1, name: "MyName" },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual("MyName");
  });

  test("can get sheet by index", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    workbook.sheets.add();
    workbook.sheets.add();
    workbook.sheets.add();
    workbook.sheets.add();

    const sheet1 = workbook.sheets.get(0);
    expect(sheet1.id).toEqual(1);
    expect(sheet1.index).toEqual(0);
    expect(sheet1.name).toEqual("Sheet1");

    const sheet3 = workbook.sheets.get(2);
    expect(sheet3.id).toEqual(3);
    expect(sheet3.index).toEqual(2);
    expect(sheet3.name).toEqual("Sheet3");
  });

  test("can get sheet by index (case: after sheet deletion)", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    workbook.sheets.add(); // Sheet2
    workbook.sheets.add(); // Sheet3
    const sheetToDelete = workbook.sheets.add(); // Sheet4
    workbook.sheets.add(); // Sheet5
    workbook.sheets.add(); // Sheet6

    sheetToDelete.delete();

    const sheet5 = workbook.sheets.get(3);
    expect(sheet5.id).toEqual(5);
    expect(sheet5.index).toEqual(3);
    expect(sheet5.name).toEqual("Sheet5");
  });

  test("can rename sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const newSheet = workbook.sheets.add("OldName");

    newSheet.name = "NewName";

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
      { id: 2, index: 1, name: "NewName" },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual("NewName");
  });

  test("can delete sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    workbook.sheets.add();
    workbook.sheets.get(0).delete();

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: "Sheet2" },
    ]);
  });

  test("renaming sheets does not break existing references", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const newSheet = workbook.sheets.add("OldName");

    const cell = workbook.cell("OldName!A1");
    cell.value = 777;

    newSheet.name = "NewName";

    expect(cell.value).toEqual(777);
    expect(cell.value).toEqual(workbook.cell("NewName!A1").value);

    cell.value = 13;
    expect(workbook.cell("NewName!A1").value).toEqual(13);
  });

  test("deleting sheets does not break existing cell references for other sheets", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    workbook.sheets.add(); // Sheet2
    workbook.sheets.add(); // Sheet3
    const cellInSheet1 = workbook.cell("Sheet1!A1");
    cellInSheet1.value = 1;
    const cellInSheet2 = workbook.cell("Sheet2!A1");
    cellInSheet2.value = 2;
    const cellInSheet3 = workbook.cell("Sheet3!A1");
    cellInSheet3.value = 3;

    workbook.sheets.get("Sheet2").delete();

    expect(cellInSheet1.value).toEqual(1);
    expect(cellInSheet3.value).toEqual(3);

    expect(() => {
      cellInSheet2.value;
    }).toThrow();
  });

  test("can access cells through specific sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    const sheet = workbook.sheets.get("Sheet1");

    sheet.cell("A3").value = 13;
    expect(workbook.cell("Sheet1!A3").value).toEqual(13);
  });
});
