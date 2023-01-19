import { IWorkbookSheets, WorkbookSheets } from "./workbookSheets";
import { WasmWorkbook } from "../__generated_pkg/equalto_wasm";
import { ICell } from "./cell";
import { parseCellReference } from "../utils";

export function newWorkbook(locale: string, timezone: string): IWorkbook {
  return new Workbook(new WasmWorkbook(locale, timezone));
}

export function loadWorkbookFromMemory(
  data: Uint8Array,
  locale: string,
  timezone: string
): IWorkbook {
  return new Workbook(WasmWorkbook.loadFromMemory(data, locale, timezone));
}

export interface IWorkbook {
  get sheets(): IWorkbookSheets;
  cell(textReference: string): ICell;
  cell(sheet: number, row: number, column: number): ICell;
}

export class Workbook implements IWorkbook {
  private readonly _wasmWorkbook: WasmWorkbook;
  private readonly _sheets: WorkbookSheets;

  constructor(wasmWorkbook: WasmWorkbook) {
    this._wasmWorkbook = wasmWorkbook;
    this._sheets = new WorkbookSheets(wasmWorkbook);
  }

  get sheets(): IWorkbookSheets {
    return this._sheets;
  }

  cell(textReference: string): ICell;
  cell(sheet: number, row: number, column: number): ICell;
  cell(
    textReferenceOrSheet: string | number,
    row?: number,
    column?: number
  ): ICell {
    if (typeof textReferenceOrSheet === "string") {
      const textReference = textReferenceOrSheet;
      const reference = parseCellReference(textReference);
      if (reference === null) {
        throw new Error(
          `Cell reference error. "${textReference}" is not valid reference.`
        );
      }
      if (reference.sheetName === undefined) {
        throw new Error(
          `Cell reference error. Sheet name is required in top-level workbook cell getter.`
        );
      }
      return this.sheets
        .get(reference.sheetName)
        .cell(reference.row, reference.column);
    } else if (
      typeof textReferenceOrSheet === "number" &&
      row !== undefined &&
      column !== undefined
    ) {
      const sheet = textReferenceOrSheet;
      return this.sheets.get(sheet).cell(row, column);
    }

    throw new Error("Function Workbook.cell received unexpected parameters.");
  }
}
