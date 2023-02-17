import { initialize } from '@equalto-software/calc';
import { readFileSync } from 'fs';

describe('loadWorkbookFromMemory', () => {
  beforeAll(async () => {
    await initialize();
  });

  test('can load XLSX from memory (from file)', async () => {
    const { loadWorkbookFromMemory } = await initialize();
    const xlsxFile = readFileSync('./api/xlsx/formats.xlsx');
    const workbook = loadWorkbookFromMemory(xlsxFile);
    expect(workbook.cell('Sheet1!A1').value).toEqual('General');
    expect(workbook.cell('Sheet1!A2').value).toEqual('Percentage');
  });

  test('throws when loaded XLSX evaluates differently', async () => {
    const { loadWorkbookFromMemory } = await initialize();
    const xlsxFile = readFileSync('./api/xlsx/XLOOKUP_with_errors.xlsx');
    expect(() => {
      loadWorkbookFromMemory(xlsxFile);
    }).toThrow(
      'EqualTo produces different results when evaluating the workbook than those already ' +
        'present in the workbook.',
    );
  });

  test('throws when loaded XLSX cannot evaluate without errors on init', async () => {
    const { loadWorkbookFromMemory } = await initialize();
    const xlsxFile = readFileSync('./api/xlsx/UNSUPPORTED_FNS_DAYS_NETWORKDAYS.xlsx');
    expect(() => {
      loadWorkbookFromMemory(xlsxFile);
    }).toThrow("Sheet1!A3 ('=_xlfn.DAYS(A2,A1)'): Invalid function: _xlfn.DAYS");
  });
});
