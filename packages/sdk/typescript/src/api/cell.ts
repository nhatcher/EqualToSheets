import dayjs from 'dayjs';
import { CalcError, ErrorKind, wrapWebAssemblyError } from 'src/errors';
import { convertDayjsUTCToSpreadsheetDate, convertSpreadsheetDateToDayjsUTC } from 'src/utils';
import { WasmWorkbook } from '../__generated_pkg/equalto_wasm';
import { ISheet, Sheet } from './sheet';
import { CellStyleManager, ICellStyle, RawCellStyle } from './style';

export interface ICell {
  /**
   * @returns Sheet that cell is in.
   */
  get sheet(): ISheet;

  /**
   * @returns Row index, count starts from 1.
   */
  get row(): number;
  /**
   * @returns Column index, count starts from 1. Column A is 1, column B is 2, ... .
   */
  get column(): number;
  /**
   * General purpose getter for cell value.
   *
   * Note: `Date` cannot be ever returned in practice due to way spreadsheets operate on dates.
   * Use `dateValue` for that purpose.
   */
  get value(): string | number | boolean | Date | null;
  /**
   * General purpose setter for cell value.
   *
   * Assigning `null` will set the cell empty without clearing it's properties like style.
   * For complete deletion see `ICell.delete()`
   *
   * Note: Dates are converted to numbers automatically. Date then can be read using `dateValue`
   * getter.
   */
  set value(value: string | number | boolean | Date | null);
  /**
   * @throws {@link CalcError} will throw if value is not valid number, or that number doesn't
   * correspond to valid date
   */
  get dateValue(): Date;
  /**
   * @throws {@link CalcError} will throw if value is not of string type
   */
  get stringValue(): string;
  /**
   * @throws {@link CalcError} will throw if value is not of number type
   */
  get numberValue(): number;
  /**
   * @throws {@link CalcError} will throw if value is not of boolean type
   */
  get booleanValue(): boolean;
  /**
   * Returns formatted cell value. Text is consistent with text displayed in cell
   * in graphical user interface.
   */
  get formattedValue(): string;
  /**
   * Returns formula if cell contains it, `null` otherwise.
   */
  get formula(): string | null;
  /**
   * Sets formula in cell, eg.: `=A2*3`
   * @throws {@link CalcError} will throw if formula is not valid
   */
  set formula(formula: string | null);
  /**
   * Sets user input in cell, eg.: `=A2*3`, `23`, `$23`, handles conversion to correct types.
   * @throws {@link CalcError} will throw if input is not valid
   */
  set input(input: string);
  /**
   * Deletes cell content along with it's properties (including style).
   */
  delete(): void;

  get style(): ICellStyle;

  /**
   * Sets the cell style to style of another cell.
   *
   * Only styles returned by cells should be passed in here.
   */
  set style(cellStyle: ICellStyle);
}

export class Cell implements ICell {
  private _wasmWorkbook: WasmWorkbook;
  private _sheet: Sheet;
  private _row: number;
  private _column: number;

  constructor(wasmWorkbook: WasmWorkbook, sheet: Sheet, row: number, column: number) {
    this._wasmWorkbook = wasmWorkbook;
    this._sheet = sheet;
    this._row = row;
    this._column = column;
  }

  get sheet(): ISheet {
    return this._sheet;
  }

  get row(): number {
    return this._row;
  }

  get column(): number {
    return this._column;
  }

  /**
   * Note: Cannot be `Date`, since XLSX stores dates as numbers. See `dateValue` for dates instead.
   */
  get value(): string | number | boolean | Date | null {
    try {
      return this._wasmWorkbook.getCellValueByIndex(this._sheet.index, this._row, this._column);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  set value(value: string | number | boolean | Date | null) {
    try {
      if (value === null || typeof value === 'string') {
        this._wasmWorkbook.updateCellWithText(
          this._sheet.index,
          this._row,
          this._column,
          value ?? '',
        );
      } else if (typeof value === 'number') {
        this._wasmWorkbook.updateCellWithNumber(this._sheet.index, this._row, this._column, value);
      } else if (typeof value === 'boolean') {
        this._wasmWorkbook.updateCellWithBool(this._sheet.index, this._row, this._column, value);
      } else if (value instanceof Date) {
        const date = dayjs.utc(value);
        const excelDate = convertDayjsUTCToSpreadsheetDate(date);
        if (excelDate < 0) {
          throw new CalcError(`Date "${date.toISOString()}" is not representable in workbook.`);
        }
        this._wasmWorkbook.updateCellWithNumber(
          this._sheet.index,
          this._row,
          this._column,
          excelDate,
        );
      }

      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get dateValue(): Date {
    const value = this.value;
    if (typeof value !== 'number') {
      throw new CalcError(
        'Cell value is not a number. Underlying number value is required for dates.',
        ErrorKind.TypeError,
      );
    }
    if (value < 0) {
      throw new CalcError(`Number "${value}" cannot be converted to date.`);
    }
    return convertSpreadsheetDateToDayjsUTC(value).toDate();
  }

  get stringValue(): string {
    const value = this.value;
    if (typeof value !== 'string') {
      throw new CalcError(
        `Type of cell's value is not string, cell value: ${JSON.stringify(value)}`,
        ErrorKind.TypeError,
      );
    }
    return value;
  }

  get numberValue(): number {
    const value = this.value;
    if (typeof value !== 'number') {
      throw new CalcError(
        `Type of cell's value is not number, cell value: ${JSON.stringify(value)}`,
        ErrorKind.TypeError,
      );
    }
    return value;
  }

  get booleanValue(): boolean {
    const value = this.value;
    if (typeof value !== 'boolean') {
      throw new CalcError(
        `Type of cell's value is not boolean, cell value: ${JSON.stringify(value)}`,
        ErrorKind.TypeError,
      );
    }
    return value;
  }

  get formattedValue(): string {
    try {
      return this._wasmWorkbook.getFormattedCellValue(this._sheet.index, this._row, this._column);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get formula(): string | null {
    try {
      return this._wasmWorkbook.getCellFormula(this._sheet.index, this._row, this._column) ?? null;
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  set formula(formula: string | null) {
    try {
      if (formula !== null) {
        this._wasmWorkbook.updateCellWithFormula(
          this._sheet.index,
          this._row,
          this._column,
          formula,
        );
      } else {
        this._wasmWorkbook.updateCellWithText(this._sheet.index, this._row, this._column, '');
      }
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  set input(input: string) {
    try {
      this._wasmWorkbook.setUserInput(this._sheet.index, this._row, this._column, input);
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  delete(): void {
    try {
      this._wasmWorkbook.deleteCell(this._sheet.index, this._row, this._column);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get style(): ICellStyle {
    try {
      const rawStyle = JSON.parse(
        this._wasmWorkbook.getCellStyle(this._sheet.index, this._row, this._column),
      ) as RawCellStyle;
      return new CellStyleManager(this._wasmWorkbook, this, rawStyle);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  set style(cellStyle: ICellStyle) {
    if (!(cellStyle instanceof CellStyleManager)) {
      throw new Error('Provided cell style is not style returned by EqualTo API.');
    }
    const sourceCell = cellStyle._getCell();
    const destinationCell = this;
    try {
      this._wasmWorkbook.copyCellStyle(
        sourceCell.sheet.index,
        sourceCell.row,
        sourceCell.column,
        destinationCell.sheet.index,
        destinationCell.row,
        destinationCell.column,
      );
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }
}
