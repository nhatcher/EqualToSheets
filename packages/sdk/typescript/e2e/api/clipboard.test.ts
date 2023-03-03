import { initialize } from '@equalto-software/calc';

describe('Clipboard', () => {
  test('getCopiedValueExtended works as expected', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const value = '=A1';
    const source = { sheet: 0, row: 1, column: 2 }; // B1
    // Copying B1 (formula: =A1) to C1
    expect(
      workbook.getCopiedValueExtended(value, 'Sheet1', source, { sheet: 0, row: 1, column: 3 }),
    ).toEqual('=B1');
    // Copying B1 (formula: =A1) to B2
    expect(
      workbook.getCopiedValueExtended(value, 'Sheet1', source, { sheet: 0, row: 2, column: 2 }),
    ).toEqual('=A2');

    workbook.sheets.add('Sheet2');

    // Note that we actually point to different sheet in these cases
    // Copying B1 (formula: =A1) to Sheet2!C1
    expect(
      workbook.getCopiedValueExtended(value, 'Sheet1', source, { sheet: 1, row: 1, column: 3 }),
    ).toEqual('=B1');
    // Copying B1 (formula: =A1) to Sheet2!B2
    expect(
      workbook.getCopiedValueExtended(value, 'Sheet1', source, { sheet: 1, row: 2, column: 2 }),
    ).toEqual('=A2');

    // Points to the initial sheet
    // Copying B1 (formula: =Sheet1!A1) to Sheet2!C1
    expect(
      workbook.getCopiedValueExtended('=Sheet1!A1', 'Sheet1', source, {
        sheet: 1,
        row: 1,
        column: 3,
      }),
    ).toEqual('=Sheet1!B1');
    // Copying B1 (formula: =Sheet1!A1) to Sheet2!B2
    expect(
      workbook.getCopiedValueExtended('=Sheet1!A1', 'Sheet1', source, {
        sheet: 1,
        row: 2,
        column: 2,
      }),
    ).toEqual('=Sheet1!A2');
  });

  test('getCutValueMoved works as expected', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const value = '=A1';
    const source = { sheet: 0, row: 1, column: 2 }; // B1
    let sourceArea = { sheet: 0, row: 1, column: 2, width: 1, height: 1 }; // B1:B1

    // Reference is not changed if outside of source area
    // Cut B1 (formula: =A1) and paste to C1
    expect(
      workbook.getCutValueMoved(value, source, { sheet: 0, row: 1, column: 3 }, sourceArea),
    ).toEqual('=A1');
    // Cut B1 (formula: =A1) and paste to B2
    expect(
      workbook.getCutValueMoved(value, source, { sheet: 0, row: 2, column: 2 }, sourceArea),
    ).toEqual('=A1');

    workbook.sheets.add('Sheet2');
    // Cut B1 (formula: =A1) and paste to Sheet2!C1
    expect(
      workbook.getCutValueMoved(value, source, { sheet: 1, row: 1, column: 3 }, sourceArea),
    ).toEqual('=Sheet1!A1');
    // Cut B1 (formula: =A1) and paste to Sheet2!B2
    expect(
      workbook.getCutValueMoved(value, source, { sheet: 1, row: 2, column: 2 }, sourceArea),
    ).toEqual('=Sheet1!A1');

    // If in source area, reference is changed
    sourceArea = { sheet: 0, row: 1, column: 1, width: 2, height: 1 }; // A1:B1
    // Cut A1:B1 (formula: =A1) and paste to Sheet2!C1
    expect(
      workbook.getCutValueMoved('=A1', source, { sheet: 1, row: 1, column: 3 }, sourceArea),
    ).toEqual('=B1');
    // Cut A1:B1 (formula: =A1) and paste to Sheet2!B2
    expect(
      workbook.getCutValueMoved('=A1', source, { sheet: 1, row: 2, column: 2 }, sourceArea),
    ).toEqual('=A2');
  });

  test('forwardReferences works as expected', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.get(0).cell('C1').formula = '=A1';
    workbook.sheets.get(0).cell('C2').formula = '=B2';
    workbook.forwardReferences(
      { sheet: 0, row: 1, column: 1, width: 2, height: 2 },
      { sheet: 0, column: 10, row: 10 },
    );
    expect(workbook.sheets.get(0).cell('C1').formula).toEqual('=J10');
    expect(workbook.sheets.get(0).cell('C2').formula).toEqual('=K11');
  });
});
