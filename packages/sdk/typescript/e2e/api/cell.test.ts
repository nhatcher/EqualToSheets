import { initialize, CalcError } from '@equalto-software/calc';
import { readFileSync } from 'fs';

describe('Workbook - Cell operations', () => {
  beforeAll(async () => {
    await initialize();
  });

  test('can read formula on empty cell - returns null', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get(0);
    expect(sheet.cell('A1').formula).toBe(null);
  });

  test('can read formula on cell with value - returns null', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get(0);
    sheet.cell('A1').value = 'Hello world';
    expect(sheet.cell('A1').formula).toBe(null);
  });

  test('can evaluate formulas in cells', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('A1').value = 13;
    sheet.cell('A2').formula = '=A1*3';
    expect(workbook.cell('Sheet1!A2').value).toEqual(13 * 3);
    expect(workbook.cell('Sheet1!A2').formula).toEqual('=A1*3');
  });

  test('cannot assign formula by value', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.get('Sheet1');

    sheet.cell('A1').value = 13;
    sheet.cell('A2').value = '=A1*3';
    expect(workbook.cell('Sheet1!A2').formula).toEqual(null);
    expect(workbook.cell('Sheet1!A2').value).toEqual('=A1*3');
  });

  test('can read formatted value from cell', async () => {
    const { loadWorkbookFromMemory } = await initialize();
    let xlsxFile = readFileSync('./api/xlsx/formats.xlsx');
    let workbook = loadWorkbookFromMemory(xlsxFile);

    expect(workbook.cell('Sheet1!B2').value).toEqual(0);
    expect(workbook.cell('Sheet1!B2').formattedValue).toEqual('0%');

    expect(workbook.cell('Sheet1!C2').value).toEqual(1);
    expect(workbook.cell('Sheet1!C2').formattedValue).toEqual('100%');

    expect(workbook.cell('Sheet1!D2').value).toEqual(2);
    expect(workbook.cell('Sheet1!D2').formattedValue).toEqual('200%');

    expect(workbook.cell('Sheet1!E2').value).toEqual(-0.12);
    expect(workbook.cell('Sheet1!E2').formattedValue).toEqual('-12%');
  });

  test('can delete cell with formatting', async () => {
    const { loadWorkbookFromMemory } = await initialize();
    const xlsxFile = readFileSync('./api/xlsx/formats.xlsx');
    const workbook = loadWorkbookFromMemory(xlsxFile);
    const cell = workbook.cell('Sheet1!C2');
    expect(cell.value).toEqual(1);
    expect(cell.formattedValue).toEqual('100%');

    cell.value = 2;
    expect(cell.formattedValue).toEqual('200%');

    cell.delete();
    expect(cell.value).toEqual('');
    expect(cell.formattedValue).toEqual('');

    cell.value = 2;
    expect(cell.value).toEqual(2);
    expect(cell.formattedValue).toEqual('2');
  });

  test('can read typed number value from cell', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');

    cell.value = 3;
    expect(cell.value).toEqual(3);
    expect(cell.numberValue).toEqual(3);

    cell.formula = '=2+2';
    expect(cell.value).toEqual(4);
    expect(cell.numberValue).toEqual(4);
  });

  test('can read date using typed number getter', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');
    cell.value = new Date('2020-01-01');
    expect(cell.numberValue).toEqual(43_831);
  });

  test('throws when non-number is read using typed number value getter', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');

    cell.value = true;
    expect(cell.value).toEqual(true);
    expect(() => cell.numberValue).toThrow(`Type of cell's value is not number, cell value: true`);

    cell.value = '3';
    expect(cell.value).toEqual('3');
    expect(() => cell.numberValue).toThrow(`Type of cell's value is not number, cell value: "3"`);
  });

  test('can read typed string value from cell', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');
    expect(cell.stringValue).toEqual('');

    cell.value = 'hello world';
    expect(cell.stringValue).toEqual('hello world');
  });

  test('throws when non-string is read using typed string value getter', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');

    cell.value = true;
    expect(cell.value).toEqual(true);
    expect(() => cell.stringValue).toThrow(`Type of cell's value is not string, cell value: true`);

    cell.value = 3;
    expect(cell.value).toEqual(3);
    expect(() => cell.stringValue).toThrow(`Type of cell's value is not string, cell value: 3`);
  });

  test('can read typed boolean value from cell', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');

    cell.value = true;
    expect(cell.booleanValue).toEqual(true);

    cell.formula = '=3<4';
    expect(cell.booleanValue).toEqual(true);
  });

  test('throws when non-boolean is read using typed boolean value getter', async () => {
    const { newWorkbook } = await initialize();
    let workbook = newWorkbook();

    let cell = workbook.cell('Sheet1!A1');

    cell.value = 'true';
    expect(cell.value).toEqual('true');
    expect(() => cell.booleanValue).toThrow(
      `Type of cell's value is not boolean, cell value: "true"`,
    );

    cell.value = 3;
    expect(cell.value).toEqual(3);
    expect(() => cell.booleanValue).toThrow(`Type of cell's value is not boolean, cell value: 3`);
  });

  test('supports setting dates', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const cell = workbook.cell('Sheet1!A1');
    cell.value = new Date('2015-02-14');
    expect(cell.value).toEqual(42049);
    expect(cell.dateValue).toBeInstanceOf(Date);
    expect(cell.dateValue).toEqual(new Date('2015-02-14T00:00:00.000Z'));
  });

  test('supports setting date-times', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const cell = workbook.cell('Sheet1!A1');
    cell.value = new Date('2015-02-14T13:30:00.000Z');
    expect(cell.value).toEqual(42049.5625);
    expect(cell.dateValue).toBeInstanceOf(Date);
    expect(cell.dateValue).toEqual(new Date('2015-02-14T13:30:00.000Z'));
  });

  test('cannot assign dates far in the past', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const cell = workbook.cell('Sheet1!A1');
    expect(() => {
      cell.value = new Date('1815-02-14');
    }).toThrow('Date "1815-02-14T00:00:00.000Z" is not representable in workbook.');
  });

  test('cannot read invalid dates - negative numbers', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const cell = workbook.cell('Sheet1!A1');
    cell.value = -1;
    expect(() => cell.dateValue).toThrow('Number "-1" cannot be converted to date.');
  });

  test('throws when values are read on cell from deleted sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.add();
    const cell = sheet.cell('A1');
    cell.value = 7;

    sheet.delete();

    const failCase = () => {
      cell.value;
    };

    expect(failCase).toThrow('Could not find sheet with sheetId=2');
    expect(failCase).toThrow(CalcError);
  });

  test('throws when values are set on cell from deleted sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.add();
    const cell = sheet.cell('A1');

    sheet.delete();

    const failCase = () => {
      cell.value = 8;
    };

    expect(failCase).toThrow('Could not find sheet with sheetId=2');
    expect(failCase).toThrow(CalcError);
  });

  test('throws when formula is read on cell from deleted sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.add();
    sheet.cell('A1').value = 7;
    const cell = sheet.cell('A2');
    cell.formula = '=A1+3';
    expect(cell.value).toEqual(10);

    sheet.delete();

    const failCase = () => {
      cell.formula;
    };

    expect(failCase).toThrow('Could not find sheet with sheetId=2');
    expect(failCase).toThrow(CalcError);
  });

  test('throws when formula is set on cell from deleted sheet', async () => {
    const { newWorkbook } = await initialize();
    const workbook = newWorkbook();
    const sheet = workbook.sheets.add();
    sheet.cell('A1').value = 7;
    const cell = sheet.cell('A2');
    sheet.delete();

    const failCase = () => {
      cell.formula = '=A1+3';
    };

    expect(failCase).toThrow('Could not find sheet with sheetId=2');
    expect(failCase).toThrow(CalcError);
  });

  describe('style', () => {
    test('can read number format', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      expect(cell.style.numberFormat).toEqual('general');
    });

    test('can set number format', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      cell.value = 7;
      cell.style.numberFormat = '0.00%';
      expect(cell.formattedValue).toEqual('700.00%');
      expect(cell.style.numberFormat).toEqual('0.00%');
    });

    test('setting number format updates own style snapshot', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      const style = cell.style;
      style.numberFormat = '0.00%';
      expect(style.numberFormat).toEqual('0.00%');
    });

    test('can set bold on cell', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      cell.style.font.bold = true;

      expect(cell.style.font.bold).toEqual(true);

      expect(cell.style.font.italics).toEqual(false);
      expect(cell.style.font.underline).toEqual(false);
      expect(cell.style.font.strikethrough).toEqual(false);
    });

    test('can set italics on cell', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      cell.style.font.italics = true;

      expect(cell.style.font.italics).toEqual(true);

      expect(cell.style.font.bold).toEqual(false);
      expect(cell.style.font.underline).toEqual(false);
      expect(cell.style.font.strikethrough).toEqual(false);
    });

    test('can set underline on cell', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      cell.style.font.underline = true;

      expect(cell.style.font.underline).toEqual(true);

      expect(cell.style.font.bold).toEqual(false);
      expect(cell.style.font.italics).toEqual(false);
      expect(cell.style.font.strikethrough).toEqual(false);
    });

    test('can set strikethrough on cell', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      cell.style.font.strikethrough = true;

      expect(cell.style.font.strikethrough).toEqual(true);

      expect(cell.style.font.bold).toEqual(false);
      expect(cell.style.font.italics).toEqual(false);
      expect(cell.style.font.underline).toEqual(false);
    });

    test('can unset toggleable font properties', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');

      expect(cell.style.font.bold).toEqual(false);
      expect(cell.style.font.italics).toEqual(false);
      expect(cell.style.font.underline).toEqual(false);
      expect(cell.style.font.strikethrough).toEqual(false);

      cell.style.bulkUpdate({
        font: {
          bold: true,
          italics: true,
          underline: true,
          strikethrough: true,
        },
      });

      expect(cell.style.font.bold).toEqual(true);
      expect(cell.style.font.italics).toEqual(true);
      expect(cell.style.font.underline).toEqual(true);
      expect(cell.style.font.strikethrough).toEqual(true);

      cell.style.font.bold = false;
      cell.style.font.italics = false;
      cell.style.font.underline = false;
      cell.style.font.strikethrough = false;

      expect(cell.style.font.bold).toEqual(false);
      expect(cell.style.font.italics).toEqual(false);
      expect(cell.style.font.underline).toEqual(false);
      expect(cell.style.font.strikethrough).toEqual(false);
    });

    test('can read default font color', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      expect(cell.style.font.color).toEqual('#000000');
    });

    test('can set font color', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      cell.style.font.color = '#ff0000';
      expect(cell.style.font.color).toEqual('#FF0000');
    });

    test('throws if set font color is invalid', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');
      expect(() => {
        cell.style.font.color = 'does not make sense';
      }).toThrow('Color "does not make sense" is not valid 3-channel hex color.');
      expect(cell.style.font.color).toEqual('#000000');
    });

    test('can set solid fill', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');

      expect(cell.style.fill.patternType).toEqual('none');
      expect(cell.style.fill.foregroundColor).toEqual('#FFFFFF');
      expect(cell.style.fill.backgroundColor).toEqual('#FFFFFF');

      cell.style.bulkUpdate({
        fill: {
          patternType: 'solid',
          foregroundColor: '#ff00ff',
        },
      });

      expect(cell.style.fill.patternType).toEqual('solid');
      expect(cell.style.fill.foregroundColor).toEqual('#FF00FF');
      expect(cell.style.fill.backgroundColor).toEqual('#FFFFFF');
    });

    test('throws if fill color is invalid', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');

      expect(() => {
        cell.style.fill.foregroundColor = '#fff';
      }).toThrow('Color "#fff" is not valid 3-channel hex color.');

      expect(() => {
        cell.style.fill.backgroundColor = '#aaa';
      }).toThrow('Color "#aaa" is not valid 3-channel hex color.');
    });

    test('can unset fill', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');

      cell.style.bulkUpdate({
        fill: {
          patternType: 'solid',
          foregroundColor: '#ff00ff',
        },
      });

      cell.style.fill.patternType = 'none';
      expect(cell.style.fill.patternType).toEqual('none');
      // saved, but user code should ignore the value:
      expect(cell.style.fill.foregroundColor).toEqual('#FF00FF');
    });

    test('bulk style update', async () => {
      const { newWorkbook } = await initialize();
      const workbook = newWorkbook();
      const cell = workbook.sheets.get(0).cell('A1');

      expect(cell.style.numberFormat).toEqual('general');
      expect(cell.style.font.bold).toEqual(false);

      cell.style.bulkUpdate({
        numberFormat: '0.00%',
        font: {
          bold: true,
          italics: true,
          underline: false,
          strikethrough: false,
        },
        fill: {
          patternType: 'solid',
          foregroundColor: '#ff00ff',
          backgroundColor: '#00ffff',
        },
      });

      expect(cell.style.numberFormat).toEqual('0.00%');
      expect(cell.style.font.bold).toEqual(true);
      expect(cell.style.font.italics).toEqual(true);
      expect(cell.style.font.underline).toEqual(false);
      expect(cell.style.font.strikethrough).toEqual(false);
      expect(cell.style.fill.patternType).toEqual('solid');
      expect(cell.style.fill.foregroundColor).toEqual('#FF00FF');
      expect(cell.style.fill.backgroundColor).toEqual('#00FFFF');
    });
  });
});
