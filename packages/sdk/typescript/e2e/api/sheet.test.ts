import { initialize, CalcError } from '@equalto-software/calc';
import type { ISheet } from '@equalto-software/calc';

const mapSheetToObject = (sheet: ISheet) => {
  return {
    id: sheet.id,
    index: sheet.index,
    name: sheet.name,
  };
};

describe('Worksheet', () => {
  beforeAll(async () => {
    await initialize();
  });

  test('can list sheets in new workbook', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: 'Sheet1' },
    ]);
  });

  test('can create sheet with default name in new workbook', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const newSheet = workbook.sheets.add();

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: 'Sheet1' },
      { id: 2, index: 1, name: 'Sheet2' },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual('Sheet2');
  });

  test('can create sheet with given name in new workbook', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const newSheet = workbook.sheets.add('MyName');

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: 'Sheet1' },
      { id: 2, index: 1, name: 'MyName' },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual('MyName');
  });

  test('can use emojis in workbook name ', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add('🙈');

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: 'Sheet1' },
      { id: 2, index: 1, name: '🙈' },
    ]);
  });

  test('can use only spaces in workbook name', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add(' ');
    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: 'Sheet1' },
      { id: 2, index: 1, name: ' ' },
    ]);
  });

  test('throws when new sheet name is blank', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();

    const failCase = () => {
      workbook.sheets.add('');
    };

    expect(failCase).toThrow("Invalid name for a sheet: ''");
    expect(failCase).toThrow(CalcError);
  });

  test('throws when new sheet name is longer than 31 characters', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();

    let name31 = 'AAAAAAAAAA' + 'BBBBBBBBBB' + 'CCCCCCCCCC' + 'D'; // len 3*10+1
    workbook.sheets.add(name31);

    const failCase = () => {
      workbook.sheets.add(name31 + 'E');
    };

    expect(failCase).toThrow("Invalid name for a sheet: 'AAAAAAAAAABBBBBBBBBBCCCCCCCCCCDE'");
    expect(failCase).toThrow(CalcError);
  });

  test('throws when name of new sheet is a duplicate', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add('MyName');

    const failCase = () => {
      workbook.sheets.add('MyName');
    };

    expect(failCase).toThrow('A worksheet already exists with that name');
    expect(failCase).toThrow(CalcError);
  });

  test('can get sheet by index', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add();
    workbook.sheets.add();
    workbook.sheets.add();
    workbook.sheets.add();

    const sheet1 = workbook.sheets.get(0);
    expect(sheet1.id).toEqual(1);
    expect(sheet1.index).toEqual(0);
    expect(sheet1.name).toEqual('Sheet1');

    const sheet3 = workbook.sheets.get(2);
    expect(sheet3.id).toEqual(3);
    expect(sheet3.index).toEqual(2);
    expect(sheet3.name).toEqual('Sheet3');
  });

  test('can get sheet by index (case: after sheet deletion)', async () => {
    const { newWorkbook } = await initialize();
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
    expect(sheet5.name).toEqual('Sheet5');
  });

  test.each<number>([-1, 10])('throws when getting sheet by invalid index (%i)', async (index) => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.sheets.get(index);
    };
    expect(failCase).toThrow(`Could not find sheet at index=${index}`);
    expect(failCase).toThrow(CalcError);
  });

  test('can get sheet by name', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add();
    workbook.sheets.add();
    workbook.sheets.add('Calculation');
    workbook.sheets.add();

    const calculationSheet = workbook.sheets.get('Calculation');
    expect(calculationSheet.id).toEqual(4);
    expect(calculationSheet.index).toEqual(3);
    expect(calculationSheet.name).toEqual('Calculation');
  });

  test('throws when getting sheet by invalid name', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.sheets.get('DoesNotExist');
    };
    expect(failCase).toThrow('Could not find sheet with name="DoesNotExist"');
    expect(failCase).toThrow(CalcError);
  });

  test('can rename sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const newSheet = workbook.sheets.add('OldName');

    newSheet.name = 'NewName';

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 1, index: 0, name: 'Sheet1' },
      { id: 2, index: 1, name: 'NewName' },
    ]);

    expect(newSheet.id).toEqual(2);
    expect(newSheet.index).toEqual(1);
    expect(newSheet.name).toEqual('NewName');
  });

  test('throws when deleted sheet is renamed', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add();
    const sheet = workbook.sheets.get(0);

    sheet.delete();

    const failCase = () => {
      sheet.name = 'hello';
    };

    expect(failCase).toThrow('Could not find sheet with sheetId=1');
    expect(failCase).toThrow(CalcError);

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: 'Sheet2' },
    ]);
  });

  test('can delete sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add();
    workbook.sheets.get(0).delete();

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: 'Sheet2' },
    ]);
  });

  test('throws when sheet is deleted multiple times', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add();
    const sheet = workbook.sheets.get(0);

    sheet.delete();

    const failCase = () => {
      sheet.delete();
    };

    expect(failCase).toThrow('Sheet not found');
    expect(failCase).toThrow(CalcError);

    expect(workbook.sheets.all().map(mapSheetToObject)).toEqual([
      { id: 2, index: 0, name: 'Sheet2' },
    ]);
  });

  test('renaming sheets does not break existing references', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const newSheet = workbook.sheets.add('OldName');

    const cell = workbook.cell('OldName!A1');
    cell.value = 777;

    newSheet.name = 'NewName';

    expect(cell.value).toEqual(777);
    expect(cell.value).toEqual(workbook.cell('NewName!A1').value);

    cell.value = 13;
    expect(workbook.cell('NewName!A1').value).toEqual(13);
  });

  test('deleting sheets does not break existing cell references for other sheets', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    workbook.sheets.add(); // Sheet2
    workbook.sheets.add(); // Sheet3
    const cellInSheet1 = workbook.cell('Sheet1!A1');
    cellInSheet1.value = 1;
    const cellInSheet2 = workbook.cell('Sheet2!A1');
    cellInSheet2.value = 2;
    const cellInSheet3 = workbook.cell('Sheet3!A1');
    cellInSheet3.value = 3;

    workbook.sheets.get('Sheet2').delete();

    expect(cellInSheet1.value).toEqual(1);
    expect(cellInSheet3.value).toEqual(3);

    expect(() => {
      cellInSheet2.value;
    }).toThrow('Could not find sheet with sheetId=2');

    expect(() => {
      cellInSheet2.value;
    }).toThrow(CalcError);
  });

  test('can access cells through specific sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('A3').value = 13;
    expect(workbook.cell('Sheet1!A3').value).toEqual(13);
  });

  test('throws when cells reference is invalid', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    const failCase = () => {
      sheet.cell('3D');
    };

    expect(failCase).toThrow('Cell reference error. "3D" is not valid reference.');
    expect(failCase).toThrow(CalcError);
  });

  test('throws when accessing cells through specific sheet with sheet specifier (Sheet!...)', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    const failCase = () => {
      sheet.cell('Sheet2!A3');
    };

    expect(failCase).toThrow(
      'Cell reference error. Sheet name cannot be specified in sheet cell getter.',
    );
    expect(failCase).toThrow(CalcError);
  });

  test('throws when cell is accessed through workbook without sheet name in reference', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.cell('A1');
    };
    expect(failCase).toThrow(
      'Cell reference error. Sheet name is required in top-level workbook cell getter.',
    );
    expect(failCase).toThrow(CalcError);
  });

  test('throws when cell is accessed by invalid sheet name in reference', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const failCase = () => {
      workbook.cell('DoesNotExist!A1');
    };
    expect(failCase).toThrow('Could not find sheet with name="DoesNotExist"');
    expect(failCase).toThrow(CalcError);
  });

  test('can set column widths', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    expect(sheet.getColumnWidth(1)).toEqual(100);
    expect(sheet.getColumnWidth(2)).toEqual(100);
    expect(sheet.getColumnWidth(3)).toEqual(100);

    sheet.setColumnWidth(2, 215);
    sheet.setColumnWidth(3, 60);

    expect(sheet.getColumnWidth(1)).toEqual(100);
    expect(sheet.getColumnWidth(2)).toEqual(215);
    expect(sheet.getColumnWidth(3)).toEqual(60);

    expect(sheet.getRowHeight(1)).toEqual(21);
    expect(sheet.getRowHeight(2)).toEqual(21);
    expect(sheet.getRowHeight(3)).toEqual(21);
  });

  test.each([-1, 0, 17_000])('throws when reading width of invalid column (%d)', async (column) => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');
    expect(() => sheet.getColumnWidth(column)).toThrow(`Column number '${column}' is not valid.`);
  });

  test.each([-1, 0, 16385])('throws when setting width of invalid column (%d)', async (column) => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');
    expect(() => sheet.setColumnWidth(column, 45)).toThrow(
      `Column number '${column}' is not valid.`,
    );
  });

  test.each([-1, 0, 1048577])('throws when reading height of invalid row (%d)', async (column) => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');
    expect(() => sheet.getRowHeight(column)).toThrow(`Row number '${column}' is not valid.`);
  });

  test.each([-1, 0, 1048577])('throws when setting height of invalid row (%d)', async (column) => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');
    expect(() => sheet.setRowHeight(column, 30)).toThrow(`Row number '${column}' is not valid.`);
  });

  test('can set row heights', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    expect(sheet.getRowHeight(1)).toEqual(21);
    expect(sheet.getRowHeight(2)).toEqual(21);
    expect(sheet.getRowHeight(3)).toEqual(21);

    sheet.setRowHeight(2, 75);
    sheet.setRowHeight(3, 18);

    expect(sheet.getRowHeight(1)).toEqual(21);
    expect(sheet.getRowHeight(2)).toEqual(75);
    expect(sheet.getRowHeight(3)).toEqual(18);

    expect(sheet.getColumnWidth(1)).toEqual(100);
    expect(sheet.getColumnWidth(2)).toEqual(100);
    expect(sheet.getColumnWidth(3)).toEqual(100);
  });

  test('can insert columns', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('A2').formula = '=A1';
    sheet.cell('B2').formula = '=B1';
    sheet.cell('C2').formula = '=A1';

    sheet.insertColumns(2, 2);

    expect(sheet.cell('A2').formula).toEqual('=A1');

    expect(sheet.cell('B2').formula).toEqual(null);
    expect(sheet.cell('C2').formula).toEqual(null);

    expect(sheet.cell('D2').formula).toEqual('=D1');
    expect(sheet.cell('E2').formula).toEqual('=A1');
  });

  test('can delete columns', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('A2').formula = '=A1';
    sheet.cell('B2').formula = '=B1';
    sheet.cell('C2').formula = '=C1';
    sheet.cell('D2').formula = '=D1';
    sheet.cell('E2').formula = '=A1';

    sheet.deleteColumns(2, 2);

    expect(sheet.cell('A2').formula).toEqual('=A1');

    expect(sheet.cell('B2').formula).toEqual('=B1');
    expect(sheet.cell('C2').formula).toEqual('=A1');

    expect(sheet.cell('D2').formula).toEqual(null);
    expect(sheet.cell('E2').formula).toEqual(null);
  });

  test('can insert rows', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('B1').formula = '=A1';
    sheet.cell('B2').formula = '=A2';
    sheet.cell('B3').formula = '=A1';

    sheet.insertRows(2, 2);

    expect(sheet.cell('B1').formula).toEqual('=A1');

    expect(sheet.cell('B2').formula).toEqual(null);
    expect(sheet.cell('B3').formula).toEqual(null);

    expect(sheet.cell('B4').formula).toEqual('=A4');
    expect(sheet.cell('B5').formula).toEqual('=A1');
  });

  test('can delete rows', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('B1').formula = '=A1';
    sheet.cell('B2').formula = '=A2';
    sheet.cell('B3').formula = '=A3';
    sheet.cell('B4').formula = '=A4';
    sheet.cell('B5').formula = '=A1';

    sheet.deleteRows(2, 2);

    expect(sheet.cell('B1').formula).toEqual('=A1');

    expect(sheet.cell('B2').formula).toEqual('=A2');
    expect(sheet.cell('B3').formula).toEqual('=A1');

    expect(sheet.cell('B4').formula).toEqual(null);
    expect(sheet.cell('B5').formula).toEqual(null);
  });

  describe('dimensions', () => {
    test('calculates dimension of empty sheet', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const sheet = workbook.sheets.get(0);
      expect(sheet.getDimensions()).toEqual({
        minRow: 1,
        maxRow: 1,
        minColumn: 1,
        maxColumn: 1,
      });
    });

    test('calculates dimension of sheet with one cell', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const sheet = workbook.sheets.get(0);
      sheet.cell('AZ187').value = 3;
      expect(sheet.getDimensions()).toEqual({
        minRow: 187,
        maxRow: 187,
        minColumn: 52,
        maxColumn: 52,
      });
    });

    test('calculates dimension of sheet with multiple cells set', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const sheet = workbook.sheets.get(0);

      sheet.cell('AZ187').value = 3;
      expect(sheet.getDimensions()).toEqual({
        minRow: 187,
        maxRow: 187,
        minColumn: 52,
        maxColumn: 52,
      });

      sheet.cell('Z15').value = 3;
      sheet.cell('AA15').value = 3;
      expect(sheet.getDimensions()).toEqual({
        minRow: 15,
        maxRow: 187,
        minColumn: 26,
        maxColumn: 52,
      });

      sheet.cell('BA18').value = 3;
      expect(sheet.getDimensions()).toEqual({
        minRow: 15,
        maxRow: 187,
        minColumn: 26,
        maxColumn: 53,
      });

      sheet.cell('A1').value = 3;
      expect(sheet.getDimensions()).toEqual({
        minRow: 1,
        maxRow: 187,
        minColumn: 1,
        maxColumn: 53,
      });

      sheet.cell('AZ188').value = 3;
      expect(sheet.getDimensions()).toEqual({
        minRow: 1,
        maxRow: 188,
        minColumn: 1,
        maxColumn: 53,
      });
    });
  });

  describe('user interface', () => {
    describe('navigateToEdgeInDirection', () => {
      const getTestWorkbook = async () => {
        const { newWorkbook } = await initialize();
        const workbook = newWorkbook();
        //   1 2 3 4 5 6
        //   A B C D E F
        // 1
        // 2     X
        // 3   X X X   X
        // 4     X
        // 5
        // 6     X
        const sheet = workbook.sheets.get(0);
        for (const cell of ['B3', 'C2', 'C3', 'C4', 'C6', 'D3', 'F3']) {
          sheet.cell(cell).value = 1;
        }
        return workbook;
      };

      test('direction = up', async () => {
        const workbook = await getTestWorkbook();
        const sheet = workbook.sheets.get(0);

        expect(sheet.userInterface.navigateToEdgeInDirection(8, 3, 'up')).toEqual([6, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(6, 3, 'up')).toEqual([4, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(4, 3, 'up')).toEqual([2, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(2, 3, 'up')).toEqual([1, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(1, 3, 'up')).toEqual([1, 3]);

        expect(sheet.userInterface.navigateToEdgeInDirection(1, 1, 'up')).toEqual([1, 1]);
        expect(sheet.userInterface.navigateToEdgeInDirection(1_000_000, 5, 'up')).toEqual([1, 5]);
      });

      test('direction = down', async () => {
        const workbook = await getTestWorkbook();
        const sheet = workbook.sheets.get(0);

        expect(sheet.userInterface.navigateToEdgeInDirection(1, 3, 'down')).toEqual([2, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(2, 3, 'down')).toEqual([4, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(4, 3, 'down')).toEqual([6, 3]);
        expect(sheet.userInterface.navigateToEdgeInDirection(6, 3, 'down')).toEqual([1_048_576, 3]);

        expect(sheet.userInterface.navigateToEdgeInDirection(1, 1, 'down')).toEqual([1_048_576, 1]);
        expect(sheet.userInterface.navigateToEdgeInDirection(1, 5, 'down')).toEqual([1_048_576, 5]);
      });

      test('direction = left', async () => {
        const workbook = await getTestWorkbook();
        const sheet = workbook.sheets.get(0);

        expect(sheet.userInterface.navigateToEdgeInDirection(3, 8, 'left')).toEqual([3, 6]);
        expect(sheet.userInterface.navigateToEdgeInDirection(3, 6, 'left')).toEqual([3, 4]);
        expect(sheet.userInterface.navigateToEdgeInDirection(3, 4, 'left')).toEqual([3, 2]);
        expect(sheet.userInterface.navigateToEdgeInDirection(3, 2, 'left')).toEqual([3, 1]);

        expect(sheet.userInterface.navigateToEdgeInDirection(1, 100, 'left')).toEqual([1, 1]);
        expect(sheet.userInterface.navigateToEdgeInDirection(5, 100, 'left')).toEqual([5, 1]);
      });

      test('direction = right', async () => {
        const workbook = await getTestWorkbook();
        const sheet = workbook.sheets.get(0);

        expect(sheet.userInterface.navigateToEdgeInDirection(3, 1, 'right')).toEqual([3, 2]);
        expect(sheet.userInterface.navigateToEdgeInDirection(3, 2, 'right')).toEqual([3, 4]);
        expect(sheet.userInterface.navigateToEdgeInDirection(3, 4, 'right')).toEqual([3, 6]);
        expect(sheet.userInterface.navigateToEdgeInDirection(3, 6, 'right')).toEqual([3, 16_384]);

        expect(sheet.userInterface.navigateToEdgeInDirection(1, 1, 'right')).toEqual([1, 16_384]);
        expect(sheet.userInterface.navigateToEdgeInDirection(5, 1, 'right')).toEqual([5, 16_384]);
      });
    });
  });
});
