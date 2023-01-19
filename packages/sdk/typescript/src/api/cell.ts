import { WasmWorkbook } from "../__generated_pkg/equalto_wasm";
import { Sheet } from "./sheet";

export interface ICell {
  get row(): number;
  get column(): number;
  get value(): string | number | boolean | Date | null;
  set value(value: string | number | boolean | Date | null);
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

  get value(): string | number | boolean | Date | null {
    return this._wasmWorkbook.getCellValueByIndex(
      this._sheet.index,
      this._row,
      this._column
    );
  }

  set value(value: string | number | boolean | Date | null) {
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
      // TODO: Support Date.
    }

    this._wasmWorkbook.evaluate();
  }

  get formula(): string | null {
    return (
      this._wasmWorkbook.getCellFormula(
        this._sheet.index,
        this._row,
        this._column
      ) ?? null
    );
  }

  set formula(formula: string | null) {
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
  }
}
