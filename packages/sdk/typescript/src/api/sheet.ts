import { ErrorKind, CalcError, wrapWebAssemblyError } from "src/errors";
import { parseCellReference } from "../utils";
import { WasmWorkbook } from "../__generated_pkg/equalto_wasm";
import { Cell, ICell } from "./cell";
import { WorkbookSheets } from "./workbookSheets";

export interface ISheet {
  /**
   * Retrieves internal ID of the worksheet. This ID is immutable.
   */
  get id(): number;

  /**
   * Retrieves position index of the worksheet. This index determines the order of tabs in graphical
   * interface. Index should be considered mutable.
   */
  get index(): number;

  /**
   * Returns worksheet name.
   */
  get name(): string;

  /**
   * Sets worksheet name. Name cannot be blank, must be at most 31 characters long and
   * cannot contain the following characters: \\ / * ? : [ ]
   * @throws {@link CalcError} will throw if name isn't valid or sheet with given name already
   * exists
   */
  set name(name: string);

  /**
   * Deletes worksheet.
   */
  delete(): void;

  /**
   * @param textReference - local cell reference, example: `A1`. It cannot include a sheet name.
   * @returns cell corresponding to provided reference.
   * @throws {@link CalcError} thrown if reference isn't valid.
   */
  cell(textReference: string): ICell;
  /**
   * @param row - row index (count starts from 1)
   * @param column - column index (count starts from 1: A=1, B=2 ...)
   * @returns cell corresponding to provided coordinates.
   * @throws {@link CalcError} thrown if reference isn't valid.
   */
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
    try {
      // TODO: Should be renamed by sheetId
      this._wasmWorkbook.renameSheetBySheetIndex(this.index, name);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
    this._workbookSheets._refreshSheetLookups();
  }

  delete(): void {
    try {
      this._wasmWorkbook.deleteSheetBySheetId(this._sheetId);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
    this._workbookSheets._refreshSheetLookups();
  }

  cell(textReference: string): ICell;
  cell(row: number, column: number): ICell;
  cell(textReferenceOrRow: string | number, column?: number): ICell {
    if (typeof textReferenceOrRow === "string") {
      const textReference = textReferenceOrRow;
      const reference = parseCellReference(textReference);
      if (reference === null) {
        throw new CalcError(
          `Cell reference error. "${textReference}" is not valid reference.`,
          ErrorKind.ReferenceError
        );
      }
      if (reference.sheetName !== undefined) {
        throw new CalcError(
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

    throw new CalcError("Function Sheet.cell received unexpected parameters.");
  }
}
