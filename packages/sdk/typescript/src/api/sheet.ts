import { ErrorKind, CalcError, wrapWebAssemblyError } from 'src/errors';
import { parseCellReference } from '../utils';
import { WasmWorkbook, WasmNavigationDirection } from '../__generated_pkg/equalto_wasm';
import { Cell, ICell } from './cell';
import { WorkbookSheets } from './workbookSheets';

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

  /**
   * @param column - column index (count starts from 1: A=1, B=2, ...)
   * @returns Column width.
   */
  getColumnWidth(column: number): number;

  /**
   * Sets column width.
   * @param column - column index (count starts from 1: A=1, B=2, ...)
   * @param columnWidth - column width
   */
  setColumnWidth(column: number, columnWidth: number): void;

  /**
   * @param row - row index (count starts from 1)
   * @returns row height
   */
  getRowHeight(row: number): number;

  /**
   * Sets row height.
   * @param row - row index (count starts from 1)
   * @param rowHeight - row height
   */
  setRowHeight(row: number, rowHeight: number): void;

  /**
   * @returns Dimensions of the worksheet.
   */
  getDimensions(): {
    minRow: number;
    maxRow: number;
    minColumn: number;
    maxColumn: number;
  };

  insertRows(row: number, rowCount: number): void;
  deleteRows(row: number, rowCount: number): void;

  insertColumns(column: number, columnCount: number): void;
  deleteColumns(column: number, columnCount: number): void;

  /**
   * Returns sub-interface providing features useful when implementing user interface
   * on top of the workbook.
   */
  get userInterface(): ISheetForUserInterface;
}

export type NavigationDirection = 'left' | 'right' | 'up' | 'down';

function navigationDirectionToWasm(direction: NavigationDirection): WasmNavigationDirection {
  switch (direction) {
    case 'left':
      return WasmNavigationDirection.Left;
    case 'right':
      return WasmNavigationDirection.Right;
    case 'up':
      return WasmNavigationDirection.Up;
    case 'down':
      return WasmNavigationDirection.Down;
    default:
      const unknownDirection: never = direction;
      throw new CalcError(`Unrecognized direction: ${unknownDirection}`);
  }
}

export interface ISheetForUserInterface {
  /**
   * @returns Cell to navigate to when CTRL+direction arrow is used.
   */
  navigateToEdgeInDirection(
    row: number,
    column: number,
    direction: NavigationDirection,
  ): [number, number];
  /**
   * @returns user input value for extend action
   * from [sourceRow, sourceColumn] to [targetRow, targetColumn]
   */
  getExtendedValue(
    sourceRow: number,
    sourceColumn: number,
    targetRow: number,
    targetColumn: number,
  ): string;
}

export class Sheet implements ISheet {
  private readonly _workbookSheets: WorkbookSheets;
  private readonly _wasmWorkbook: WasmWorkbook;
  private readonly _sheetId: number;
  private readonly _sheetForUserInterface: SheetForUserInterface;

  constructor(workbookSheets: WorkbookSheets, wasmWorkbook: WasmWorkbook, sheetId: number) {
    this._workbookSheets = workbookSheets;
    this._wasmWorkbook = wasmWorkbook;
    this._sheetId = sheetId;
    this._sheetForUserInterface = new SheetForUserInterface(this, wasmWorkbook);
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
    if (typeof textReferenceOrRow === 'string') {
      const textReference = textReferenceOrRow;
      const reference = parseCellReference(textReference);
      if (reference === null) {
        throw new CalcError(
          `Cell reference error. "${textReference}" is not valid reference.`,
          ErrorKind.ReferenceError,
        );
      }
      if (reference.sheetName !== undefined) {
        throw new CalcError(
          `Cell reference error. Sheet name cannot be specified in sheet cell getter.`,
          ErrorKind.ReferenceError,
        );
      }
      return this.cell(reference.row, reference.column);
    }

    if (typeof textReferenceOrRow === 'number' && typeof column === 'number') {
      const row = textReferenceOrRow;
      return new Cell(this._wasmWorkbook, this, row, column);
    }

    throw new CalcError('Function Sheet.cell received unexpected parameters.');
  }

  getColumnWidth(column: number): number {
    try {
      return this._wasmWorkbook.getColumnWidth(this.index, column);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  setColumnWidth(column: number, columnWidth: number): void {
    try {
      this._wasmWorkbook.setColumnWidth(this.index, column, columnWidth);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  getRowHeight(row: number): number {
    try {
      return this._wasmWorkbook.getRowHeight(this.index, row);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  setRowHeight(row: number, rowHeight: number): void {
    try {
      this._wasmWorkbook.setRowHeight(this.index, row, rowHeight);
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  getDimensions(): {
    minRow: number;
    maxRow: number;
    minColumn: number;
    maxColumn: number;
  } {
    try {
      const wasmSheetDimensions = this._wasmWorkbook.getSheetDimensions(this.index);
      const sheetDimensions = {
        minRow: wasmSheetDimensions.minRow,
        maxRow: wasmSheetDimensions.maxRow,
        minColumn: wasmSheetDimensions.minColumn,
        maxColumn: wasmSheetDimensions.maxColumn,
      };
      wasmSheetDimensions.free();
      return sheetDimensions;
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  insertRows(row: number, rowCount: number): void {
    try {
      this._wasmWorkbook.insertRows(this.index, row, rowCount);
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  deleteRows(row: number, rowCount: number): void {
    try {
      this._wasmWorkbook.deleteRows(this.index, row, rowCount);
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  insertColumns(column: number, columnCount: number): void {
    try {
      this._wasmWorkbook.insertColumns(this.index, column, columnCount);
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  deleteColumns(column: number, columnCount: number): void {
    try {
      this._wasmWorkbook.deleteColumns(this.index, column, columnCount);
      this._wasmWorkbook.evaluate();
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  get userInterface(): ISheetForUserInterface {
    return this._sheetForUserInterface;
  }
}

export class SheetForUserInterface implements ISheetForUserInterface {
  private readonly _sheet: Sheet;
  private readonly _wasmWorkbook: WasmWorkbook;

  constructor(sheet: Sheet, wasmWorkbook: WasmWorkbook) {
    this._sheet = sheet;
    this._wasmWorkbook = wasmWorkbook;
  }

  navigateToEdgeInDirection(
    row: number,
    column: number,
    direction: NavigationDirection,
  ): [number, number] {
    try {
      const cell = this._wasmWorkbook.navigateToEdgeInDirection(
        this._sheet.index,
        row,
        column,
        navigationDirectionToWasm(direction),
      );

      return [cell.row, cell.column];
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }

  getExtendedValue(
    sourceRow: number,
    sourceColumn: number,
    targetRow: number,
    targetColumn: number,
  ): string {
    try {
      return this._wasmWorkbook.getExtendedValue(
        this._sheet.index,
        sourceRow,
        sourceColumn,
        targetRow,
        targetColumn,
      );
    } catch (error) {
      throw wrapWebAssemblyError(error);
    }
  }
}
