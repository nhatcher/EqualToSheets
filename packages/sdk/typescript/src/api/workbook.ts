import { IWorkbookSheets, WorkbookSheets } from './workbookSheets';
import { WasmWorkbook, WasmArea, WasmCellReferenceIndex } from '../__generated_pkg/equalto_wasm';
import { ICell } from './cell';
import { parseCellReference } from '../utils';
import { ErrorKind, CalcError, wrapWebAssemblyError } from 'src/errors';

export function newWorkbook(): IWorkbook {
  let wasmWorkbook;
  try {
    wasmWorkbook = new WasmWorkbook('en', 'UTC');
  } catch (error) {
    throw wrapWebAssemblyError(error);
  }

  return new Workbook(wasmWorkbook);
}

export function loadWorkbookFromMemory(data: Uint8Array): IWorkbook {
  let wasmWorkbook;
  try {
    wasmWorkbook = WasmWorkbook.loadFromMemory(data, 'en', 'UTC');
  } catch (error) {
    throw wrapWebAssemblyError(error);
  }

  return new Workbook(wasmWorkbook);
}

export function loadWorkbookFromJson(workbookJson: string): IWorkbook {
  let wasmWorkbook;
  try {
    wasmWorkbook = WasmWorkbook.loadFromJson(workbookJson);
  } catch (error) {
    throw wrapWebAssemblyError(error);
  }

  return new Workbook(wasmWorkbook);
}

type CellReference = {
  sheet: number;
  row: number;
  column: number;
};

type Area = {
  sheet: number;
  row: number;
  column: number;
  width: number;
  height: number;
};

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
  /**
   * @returns Uint8Buffer containing XLSX data.
   */
  saveToXlsx(): Uint8Array;
  /**
   * @returns string with json representation of the workbook.
   */
  toJson(): string;
  /**
   * Used for getting target value if value was copied.
   * All references are extended from source to target.
   * @param value - copied value.
   * @param source_sheet_name - sheet name of the source cell
   * @param source - cell reference of source cell
   * @param target - cell reference of target cell
   * @returns Copied value extended to a target cell in a form of user input. Use cell.input to set.
   */
  getCopiedValueExtended(
    value: string,
    source_sheet_name: string,
    source: CellReference,
    target: CellReference,
  ): string;
  /**
   * Used for getting target value if value was cut.
   * If formula is referencing source_area, it needs to be moved to target_area.
   * Other references stay they same.
   * @param value - cut value.
   * @param source - cell reference of source cell.
   * @param target - cell reference of target cell.
   * @param source_area - source area - we don't move references to given area
   * @returns Cut value moved to a target cell in a form of user input. Use cell.input to set.
   */
  getCutValueMoved(
    value: string,
    source: CellReference,
    target: CellReference,
    source_area: Area,
  ): string;

  /**
   * Used for forwarding references from cut source_area to pasted target.
   * @param source_area - area that was cut.
   * @param target - target cell where it was pasted.
   */
  forwardReferences(source_area: Area, target: CellReference): void;
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
  cell(textReferenceOrSheet: string | number, row?: number, column?: number): ICell {
    if (typeof textReferenceOrSheet === 'string') {
      const textReference = textReferenceOrSheet;
      const reference = parseCellReference(textReference);
      if (reference === null) {
        throw new CalcError(
          `Cell reference error. "${textReference}" is not valid reference.`,
          ErrorKind.ReferenceError,
        );
      }
      if (reference.sheetName === undefined) {
        throw new CalcError(
          `Cell reference error. Sheet name is required in top-level workbook cell getter.`,
          ErrorKind.ReferenceError,
        );
      }
      return this.sheets.get(reference.sheetName).cell(reference.row, reference.column);
    } else if (
      typeof textReferenceOrSheet === 'number' &&
      row !== undefined &&
      column !== undefined
    ) {
      const sheet = textReferenceOrSheet;
      return this.sheets.get(sheet).cell(row, column);
    }

    throw new CalcError('Function Workbook.cell received unexpected parameters.');
  }

  saveToXlsx(): Uint8Array {
    return this._wasmWorkbook.saveToMemory();
  }

  toJson(): string {
    return this._wasmWorkbook.toJson();
  }

  getCopiedValueExtended(
    value: string,
    source_sheet_name: string,
    source: CellReference,
    target: CellReference,
  ): string {
    try {
      return this._wasmWorkbook.getCopiedValueExtended(
        value,
        source_sheet_name,
        Workbook.cellReferenceToWasm(source),
        Workbook.cellReferenceToWasm(target),
      );
    } catch (e) {
      throw wrapWebAssemblyError(e);
    }
  }

  getCutValueMoved(
    value: string,
    source: CellReference,
    target: CellReference,
    sourceArea: Area,
  ): string {
    try {
      return this._wasmWorkbook.getCutValueMoved(
        value,
        Workbook.cellReferenceToWasm(source),
        Workbook.cellReferenceToWasm(target),
        Workbook.areaToWasm(sourceArea),
      );
    } catch (e) {
      throw wrapWebAssemblyError(e);
    }
  }

  forwardReferences(sourceArea: Area, target: CellReference): void {
    try {
      this._wasmWorkbook.forwardReferences(
        Workbook.areaToWasm(sourceArea),
        Workbook.cellReferenceToWasm(target),
      );
      this._wasmWorkbook.evaluate();
    } catch (e) {
      throw wrapWebAssemblyError(e);
    }
  }

  private static cellReferenceToWasm(cell: CellReference) {
    return new WasmCellReferenceIndex(cell.sheet, cell.row, cell.column);
  }

  private static areaToWasm(area: Area) {
    return new WasmArea(area.sheet, area.row, area.column, area.width, area.height);
  }
}
