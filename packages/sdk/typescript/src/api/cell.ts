import dayjs from "dayjs";
import { ErrorKind, SheetsError, wrapWebAssemblyError } from "src/errors";
import {
  convertDayjsUTCToSpreadsheetDate,
  convertSpreadsheetDateToDayjsUTC,
} from "src/utils";
import { WasmWorkbook } from "../__generated_pkg/equalto_wasm";
import { Sheet } from "./sheet";

export interface ICell {
  get row(): number;
  get column(): number;
  get value(): string | number | boolean | Date | null;
  set value(value: string | number | boolean | Date | null);
  get dateValue(): Date;
  get stringValue(): string;
  get numberValue(): number;
  get booleanValue(): boolean;
  get formattedValue(): string;
  get formula(): string | null;
  set formula(formula: string | null);
}

export class Cell implements ICell {
  private _wasmWorkbook: WasmWorkbook;
  private _sheet: Sheet;
  private _row: number;
  private _column: number;

  constructor(
    wasmWorkbook: WasmWorkbook,
    sheet: Sheet,
    row: number,
    column: number
  ) {
    this._wasmWorkbook = wasmWorkbook;
    this._sheet = sheet;
    this._row = row;
    this._column = column;
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
      return this._wasmWorkbook.getCellValueByIndex(
        this._sheet.index,
        this._row,
        this._column
      );
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  set value(value: string | number | boolean | Date | null) {
    try {
      if (value === null || typeof value === "string") {
        this._wasmWorkbook.updateCellWithText(
          this._sheet.index,
          this._row,
          this._column,
          value ?? ""
        );
      } else if (typeof value === "number") {
        this._wasmWorkbook.updateCellWithNumber(
          this._sheet.index,
          this._row,
          this._column,
          value
        );
      } else if (typeof value === "boolean") {
        this._wasmWorkbook.updateCellWithBool(
          this._sheet.index,
          this._row,
          this._column,
          value
        );
      } else if (value instanceof Date) {
        const date = dayjs.utc(value);
        const excelDate = convertDayjsUTCToSpreadsheetDate(date);
        if (excelDate < 0) {
          throw new SheetsError(
            `Date "${date.toISOString()}" is not representable in workbook.`
          );
        }
        this._wasmWorkbook.updateCellWithNumber(
          this._sheet.index,
          this._row,
          this._column,
          excelDate
        );
      }

      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get dateValue(): Date {
    const value = this.value;
    if (typeof value !== "number") {
      throw new SheetsError(
        "Cell value is not a number. Underlying number value is required for dates.",
        ErrorKind.TypeError
      );
    }
    if (value < 0) {
      throw new SheetsError(`Number "${value}" cannot be converted to date.`);
    }
    return convertSpreadsheetDateToDayjsUTC(value).toDate();
  }

  get stringValue(): string {
    const value = this.value;
    if (typeof value !== "string") {
      throw new SheetsError(
        `Type of cell's value is not string, cell value: ${JSON.stringify(
          value
        )}`,
        ErrorKind.TypeError
      );
    }
    return value;
  }

  get numberValue(): number {
    const value = this.value;
    if (typeof value !== "number") {
      throw new SheetsError(
        `Type of cell's value is not number, cell value: ${JSON.stringify(
          value
        )}`,
        ErrorKind.TypeError
      );
    }
    return value;
  }

  get booleanValue(): boolean {
    const value = this.value;
    if (typeof value !== "boolean") {
      throw new SheetsError(
        `Type of cell's value is not boolean, cell value: ${JSON.stringify(
          value
        )}`,
        ErrorKind.TypeError
      );
    }
    return value;
  }

  get formattedValue(): string {
    try {
      return this._wasmWorkbook.getFormattedCellValue(
        this._sheet.index,
        this._row,
        this._column
      );
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get formula(): string | null {
    try {
      return (
        this._wasmWorkbook.getCellFormula(
          this._sheet.index,
          this._row,
          this._column
        ) ?? null
      );
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
          formula
        );
      } else {
        this._wasmWorkbook.updateCellWithText(
          this._sheet.index,
          this._row,
          this._column,
          ""
        );
      }
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }
}
