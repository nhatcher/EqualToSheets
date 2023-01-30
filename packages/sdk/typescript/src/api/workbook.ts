import { IWorkbookSheets, WorkbookSheets } from './workbookSheets';
import { WasmWorkbook } from '../__generated_pkg/equalto_wasm';
import { ICell } from './cell';
import { parseCellReference } from '../utils';
import { ErrorKind, CalcError, wrapWebAssemblyError } from 'src/errors';

export function newWorkbook(): IWorkbook {
  let wasmWorkbook;
  try {
    wasmWorkbook = new WasmWorkbook("en", "UTC");
  } catch (error) {
    throw wrapWebAssemblyError(error);
  }

  return new Workbook(wasmWorkbook);
}

export function loadWorkbookFromMemory(data: Uint8Array): IWorkbook {
  let wasmWorkbook;
  try {
    wasmWorkbook = WasmWorkbook.loadFromMemory(data, "en", "UTC");
  } catch (error) {
    throw wrapWebAssemblyError(error);
  }

  return new Workbook(wasmWorkbook);
}

export interface IWorkbook {
  get sheets(): IWorkbookSheets;
  /**
   * @param textReference - global cell reference, example: `Sheet1!A1`. It must include a sheet name.
   * @returns cell corresponding to provided reference.
   * @throws {@link CalcError} thrown if reference isn't valid.
   */
  cell(textReference: string): ICell;
  /**
   * @param sheet - sheet index (count starts from 0).
   * @param row - row index (count starts from 1)
   * @param column - column index (count starts from 1: A=1, B=2 ...)
   * @returns cell corresponding to provided coordinates.
   * @throws {@link CalcError} thrown if reference isn't valid.
   */
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
        throw new CalcError(
          `Cell reference error. "${textReference}" is not valid reference.`,
          ErrorKind.ReferenceError
        );
      }
      if (reference.sheetName === undefined) {
        throw new CalcError(
          `Cell reference error. Sheet name is required in top-level workbook cell getter.`,
          ErrorKind.ReferenceError
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

    throw new CalcError(
      "Function Workbook.cell received unexpected parameters."
    );
  }
}
