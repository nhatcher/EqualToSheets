import { initialize, getApi, SheetsError } from "@equalto/sheets";
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
    const workbook = newWorkbook();
    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
    ]);
  });

  test("can create sheet with default name in new workbook", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
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
    const workbook = newWorkbook();
    const newSheet = workbook.sheets.add("MyName");

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
      { id: 2, index: 1, name: "MyName" },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual("MyName");
  });

  test("can use emojis in workbook name ", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add("🙈");

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
      { id: 2, index: 1, name: "🙈" },
    ]);
  });

  test("can use only spaces in workbook name", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add(" ");
    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: "Sheet1" },
      { id: 2, index: 1, name: " " },
    ]);
  });

  test("throws when new sheet name is blank", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();

    const failCase = () => {
      workbook.sheets.add("");
    };

    expect(failCase).toThrow("Invalid name for a sheet: ''");
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when new sheet name is longer than 31 characters", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();

    let name31 = "AAAAAAAAAA" + "BBBBBBBBBB" + "CCCCCCCCCC" + "D"; // len 3*10+1
    workbook.sheets.add(name31);

    const failCase = () => {
      workbook.sheets.add(name31 + "E");
    };

    expect(failCase).toThrow(
      "Invalid name for a sheet: 'AAAAAAAAAABBBBBBBBBBCCCCCCCCCCDE'"
    );
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when name of new sheet is a duplicate", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add("MyName");

    const failCase = () => {
      workbook.sheets.add("MyName");
    };

    expect(failCase).toThrow("A worksheet already exists with that name");
    expect(failCase).toThrow(SheetsError);
  });

  test("can get sheet by index", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
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
    const workbook = newWorkbook();
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

  test.each<number>([-1, 10])(
    "throws when getting sheet by invalid index (%i)",
    async (index) => {
      const { newWorkbook } = await getApi();
      const workbook = newWorkbook();
      const failCase = () => {
        workbook.sheets.get(index);
      };
      expect(failCase).toThrow(`Could not find sheet at index=${index}`);
      expect(failCase).toThrow(SheetsError);
    }
  );

  test("can get sheet by name", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add();
    workbook.sheets.add();
    workbook.sheets.add("Calculation");
    workbook.sheets.add();

    const calculationSheet = workbook.sheets.get("Calculation");
    expect(calculationSheet.id).toEqual(4);
    expect(calculationSheet.index).toEqual(3);
    expect(calculationSheet.name).toEqual("Calculation");
  });

  test("throws when getting sheet by invalid name", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.sheets.get("DoesNotExist");
    };
    expect(failCase).toThrow('Could not find sheet with name="DoesNotExist"');
    expect(failCase).toThrow(SheetsError);
  });

  test("can rename sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
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

  test("throws when deleted sheet is renamed", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add();
    const sheet = workbook.sheets.get(0);

    sheet.delete();

    const failCase = () => {
      sheet.name = "hello";
    };

    expect(failCase).toThrow("Could not find sheet with sheetId=1");
    expect(failCase).toThrow(SheetsError);

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: "Sheet2" },
    ]);
  });

  test("can delete sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add();
    workbook.sheets.get(0).delete();

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: "Sheet2" },
    ]);
  });

  test("throws when sheet is deleted multiple times", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    workbook.sheets.add();
    const sheet = workbook.sheets.get(0);

    sheet.delete();

    const failCase = () => {
      sheet.delete();
    };

    expect(failCase).toThrow("Sheet not found");
    expect(failCase).toThrow(SheetsError);

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: "Sheet2" },
    ]);
  });

  test("renaming sheets does not break existing references", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
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
    const workbook = newWorkbook();
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
    }).toThrow("Could not find sheet with sheetId=2");

    expect(() => {
      cellInSheet2.value;
    }).toThrow(SheetsError);
  });

  test("can access cells through specific sheet", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get("Sheet1");

    sheet.cell("A3").value = 13;
    expect(workbook.cell("Sheet1!A3").value).toEqual(13);
  });

  test("throws when cells reference is invalid", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get("Sheet1");

    const failCase = () => {
      sheet.cell("3D");
    };

    expect(failCase).toThrow(
      'Cell reference error. "3D" is not valid reference.'
    );
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when accessing cells through specific sheet with sheet specifier (Sheet!...)", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get("Sheet1");

    const failCase = () => {
      sheet.cell("Sheet2!A3");
    };

    expect(failCase).toThrow(
      "Cell reference error. Sheet name cannot be specified in sheet cell getter."
    );
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when cell is accessed through workbook without sheet name in reference", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.cell("A1");
    };
    expect(failCase).toThrow(
      "Cell reference error. Sheet name is required in top-level workbook cell getter."
    );
    expect(failCase).toThrow(SheetsError);
  });

  test("throws when cell is accessed by invalid sheet name in reference", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.cell("DoesNotExist!A1");
    };
    expect(failCase).toThrow('Could not find sheet with name="DoesNotExist"');
    expect(failCase).toThrow(SheetsError);
  });
});
