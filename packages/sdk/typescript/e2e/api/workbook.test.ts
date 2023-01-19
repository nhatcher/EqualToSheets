import { initialize, getApi } from "@equalto/sheets";

describe("workbook", () => {
  beforeAll(async () => {
    await initialize();
  });

  test("can list sheets in new workbook", async () => {
    const { newWorkbook } = await getApi();
    const workbook = newWorkbook("en", "Europe/Berlin");
    expect(
      workbook.sheets.all().map((sheet) => ({
        id: sheet.id,
        index: sheet.index,
        name: sheet.name,
      }))
    ).toEqual([{ id: 1, index: 0, name: "Sheet1" }]);
  });
});
