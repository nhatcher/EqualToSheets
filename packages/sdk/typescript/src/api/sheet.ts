import { ErrorKind, SheetsError, wrapWebAssemblyCall } from "src/errors";
import { parseCellReference } from "../utils";
import { WasmWorkbook } from "../__generated_pkg/equalto_wasm";
import { Cell, ICell } from "./cell";
import { WorkbookSheets } from "./workbookSheets";

export interface ISheet {
  get id(): number;
  get index(): number;

  get name(): string;
  set name(name: string);

  delete(): void;

  cell(textReference: string): ICell;
  cell(row: number, column: number): ICell;
}

export class Sheet implements ISheet {
  private readonly _workbookSheets: WorkbookSheets;
  private readonly _wasmWorkbook: WasmWorkbook;
  private readonly _sheetId: number;

  constructor(
    workbookSheets: WorkbookSheets,
    wasmWorkbook: WasmWorkbook,
    sheetId: number
  ) {
    this._workbookSheets = workbookSheets;
    this._wasmWorkbook = wasmWorkbook;
    this._sheetId = sheetId;
  }

  get id(): number {
    return this._sheetId;
  }

  get index(): number {
    return this._workbookSheets._getSheetIndexBySheetId(this._sheetId);
  }

  get name(): string {
    return this._workbookSheets._getSheetNameBySheetId(this._sheetId);
  }

  set name(name: string) {
    wrapWebAssemblyCall(() => {
      // TODO: Should be renamed by sheetId
      this._wasmWorkbook.renameSheetBySheetIndex(this.index, name);
    });
    this._workbookSheets._refreshSheetLookups();
  }

  delete(): void {
    wrapWebAssemblyCall(() => {
      this._wasmWorkbook.deleteSheetBySheetId(this._sheetId);
    });
    this._workbookSheets._refreshSheetLookups();
  }

  cell(textReference: string): ICell;
  cell(row: number, column: number): ICell;
  cell(textReferenceOrRow: string | number, column?: number): ICell {
    if (typeof textReferenceOrRow === "string") {
      const textReference = textReferenceOrRow;
      const reference = parseCellReference(textReference);
      if (reference === null) {
        throw new SheetsError(
          `Cell reference error. "${textReference}" is not valid reference.`,
          ErrorKind.ReferenceError
        );
      }
      if (reference.sheetName !== undefined) {
        throw new SheetsError(
          `Cell reference error. Sheet name cannot be specified in sheet cell getter.`,
          ErrorKind.ReferenceError
        );
      }
      return this.cell(reference.row, reference.column);
    }

    if (typeof textReferenceOrRow === "number" && typeof column === "number") {
      const row = textReferenceOrRow;
      return new Cell(this._wasmWorkbook, this, row, column);
    }

    throw new SheetsError(
      "Function Sheet.cell received unexpected parameters."
    );
  }
}
