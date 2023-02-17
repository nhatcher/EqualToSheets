import { initialize } from '@equalto-software/calc';
import { writeFileSync, unlinkSync, existsSync, readFileSync } from 'fs';

const TEST_XLSX_FILE = '_TEST_XLSX_FILE_.xlsx';

describe('XLSX import/export', () => {
  afterEach(() => {
    if (existsSync(TEST_XLSX_FILE)) {
      unlinkSync(TEST_XLSX_FILE);
    }
  });

  test('save simple workbook and then load it immediately', async () => {
    const { newWorkbook, loadWorkbookFromMemory } = await initialize();

    {
      const workbookToSave = newWorkbook();
      const sheet = workbookToSave.sheets.get(0);
      sheet.cell('A1').value = 1;
      sheet.cell('B1').value = 2;
      sheet.cell('C1').formula = '=A1+B1';
      const bufferToSave = workbookToSave.saveToXlsx();
      writeFileSync(TEST_XLSX_FILE, bufferToSave);
    }

    {
      const bufferToLoad = readFileSync(TEST_XLSX_FILE);
      const loadedWorkbook = loadWorkbookFromMemory(bufferToLoad);
      expect(loadedWorkbook.cell('Sheet1!A1').value).toEqual(1);
      expect(loadedWorkbook.cell('Sheet1!B1').value).toEqual(2);
      expect(loadedWorkbook.cell('Sheet1!C1').value).toEqual(3);
      expect(loadedWorkbook.cell('Sheet1!C1').formula).toEqual('=A1+B1');
    }
  });
});
