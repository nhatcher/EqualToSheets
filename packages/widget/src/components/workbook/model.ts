/* eslint-disable class-methods-use-this */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { FormulaToken, ICell, IWorkbook, NavigationDirection } from '@equalto-software/calc';
import Papa from 'papaparse';
import { TabsInput } from './components/navigation/common';
import { workbookLastColumn, workbookLastRow } from './constants';
import { Area, Cell, NavigationKey } from './util';

export enum ValueType {
  Boolean = 'boolean',
  Number = 'number',
  Text = 'text',
}

export type ExcelValue = string | number | boolean;

interface ModelSettings {
  workbook: IWorkbook;
  getTokens: (formula: string) => FormulaToken[];
}

interface RowDiff {
  rowHeight: number;
  rowData: string;
}

interface SetColumnWidthDiff {
  type: 'set_column_width';
  sheet: number;
  column: number;
  newValue: number;
  oldValue: number;
}

interface SetRowHeightDiff {
  type: 'set_row_height';
  sheet: number;
  row: number;
  newValue: number;
  oldValue: number;
}

export type CellValue = string | number | boolean | Date | null;

interface SetCellValueDiff {
  type: 'set_cell_value';
  sheet: number;
  column: number;
  row: number;
  newValue: CellValue;
  newStyle: number;
  oldValue: CellValue;
  oldStyle: number;
}

interface DeleteCellDiff {
  type: 'delete_cell';
  sheet: number;
  column: number;
  row: number;
  oldValue: CellValue;
  oldStyle: number;
}

interface RemoveCellDiff {
  type: 'remove_cell';
  sheet: number;
  column: number;
  row: number;
  oldValue: CellValue;
  oldStyle: number;
}

interface SetCellStyleDiff {
  type: 'set_cell_style';
  sheet: number;
  column: number;
  row: number;
  oldValue: CellStyle;
  newValue: CellStyle;
}

interface InsertRowDiff {
  type: 'insert_row';
  sheet: number;
  row: number;
}

interface DeleteRowDiff {
  type: 'delete_row';
  sheet: number;
  row: number;
  oldValue: RowDiff;
}

type Diff =
  | SetCellValueDiff
  | SetColumnWidthDiff
  | SetRowHeightDiff
  | DeleteCellDiff
  | RemoveCellDiff
  | SetCellStyleDiff
  | InsertRowDiff
  | DeleteRowDiff;

function assertUnreachable(value: never): never {
  throw new Error(`Unreachable value ${value}`);
}

interface SheetArea extends Area {
  sheet: number;
}

export type CellStyle = ICell['style'];

interface CellData {
  value: string | null;
  style: ICell['style'] | null;
}

interface RowData {
  [column: number]: CellData;
}

interface SheetData {
  sheetName: string;
  data: {
    [row: number]: RowData;
  };
}

type StyleReducer = (style: ICell['style']) => Parameters<ICell['style']['bulkUpdate']>[0];

type PasteType = 'copy' | 'cut';

// Replaces all tabs with spaces in a string
function escapeTabs(s: string): string {
  return s.replace(/\t/g, '  ');
}

function getValueFromSheetData(sheetData: SheetData, row: number, column: number): CellData {
  const rowData = sheetData.data[row];
  if (!rowData) {
    return {
      style: null,
      value: '',
    };
  }
  const cellData = rowData[column];
  if (!cellData) {
    return {
      style: null,
      value: '',
    };
  }
  return cellData;
}

function isCellInArea(sheet: number, row: number, column: number, area: SheetArea): boolean {
  if (sheet !== area.sheet) {
    return false;
  }
  if (row < area.rowStart || row > area.rowEnd) {
    return false;
  }
  if (column < area.columnStart || column > area.columnEnd) {
    return false;
  }
  return true;
}

export type Change = {
  type: string;
};
type Subscriber = (change: Change) => void;

export default class Model {
  getTokens: (formula: string) => FormulaToken[];

  private history: { undo: Diff[][]; redo: Diff[][] };

  private workbook: IWorkbook;

  private nextSubscriberKey = 0;

  private subscribers: Record<number, Subscriber> = {};

  constructor(options: ModelSettings) {
    this.workbook = options.workbook;
    this.getTokens = options.getTokens;
    this.history = {
      undo: [],
      redo: [],
    };
  }

  subscribe(subscriber: Subscriber): number {
    const key = this.nextSubscriberKey;
    this.nextSubscriberKey += 1;
    this.subscribers[key] = subscriber;
    return key;
  }

  unsubscribe(key: number): void {
    if (key in this.subscribers) {
      delete this.subscribers[key];
    }
  }

  private notifySubscribers(change: Change) {
    for (const subscriber of Object.values(this.subscribers)) {
      subscriber(change);
    }
  }

  setSheetColor(sheet: number, color: string): void {
    // TODO: this.wasm.set_sheet_color(sheet, color);
    this.notifySubscribers({ type: 'setSheetColor' });
  }

  getColumnWidth(sheet: number, column: number): number {
    return this.workbook.sheets.get(sheet).getColumnWidth(column);
  }

  setColumnWidth(sheet: number, column: number, width: number): void {
    this.workbook.sheets.get(sheet).setColumnWidth(column, width);
    this.notifySubscribers({ type: 'setColumnWidth' });
  }

  getTabs(): TabsInput[] {
    return this.workbook.sheets.all().map((sheet) => ({
      index: sheet.index,
      sheet_id: sheet.id,
      name: sheet.name,
      state: 'visible',
    }));
  }

  getRowHeight(sheet: number, row: number): number {
    return this.workbook.sheets.get(sheet).getRowHeight(row);
  }

  setRowHeight(sheet: number, row: number, height: number): void {
    this.workbook.sheets.get(sheet).setRowHeight(row, height);
    this.notifySubscribers({ type: 'setRowHeight' });
  }

  getTextAt(sheet: number, row: number, column: number): string {
    const cell = this.workbook.cell(sheet, row, column);
    return cell.formula ?? `${cell.value ?? ''}`;
  }

  getUICell(sheet: number, row: number, column: number): ICell {
    return this.workbook.cell(sheet, row, column);
  }

  getSheetDimensions(sheet: number) {
    return this.workbook.sheets.get(sheet).getDimensions();
  }

  getCellStyle(sheet: number, row: number, column: number): ICell['style'] {
    return this.workbook.sheets.get(sheet).cell(row, column).style;
  }

  addBlankSheet(): void {
    this.workbook.sheets.add();
    this.notifySubscribers({ type: 'addBlankSheet' });
  }

  renameSheet(sheet: number, newName: string): void {
    this.workbook.sheets.get(sheet).name = newName;
    this.notifySubscribers({ type: 'renameSheet' });
  }

  deleteSheet(sheet: number): void {
    this.workbook.sheets.get(sheet).delete();
    this.notifySubscribers({ type: 'deleteSheet' });
  }

  setCellsStyle(sheet: number, area: Area, reducer: StyleReducer): void {
    let { rowStart, rowEnd, columnStart, columnEnd } = area;
    if (rowStart > rowEnd) {
      [rowStart, rowEnd] = [rowEnd, rowStart];
    }
    if (columnStart > columnEnd) {
      [columnStart, columnEnd] = [columnEnd, columnStart];
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const cell = this.workbook.sheets.get(sheet).cell(row, column);
        cell.style.bulkUpdate(reducer(cell.style));
      }
    }
    this.notifySubscribers({ type: 'setCellsStyle' });
  }

  setNumberFormat(sheet: number, area: Area, numberFormat: string): void {
    this.setCellsStyle(sheet, area, () => ({ numberFormat }));
  }

  setFillColor(sheet: number, area: Area, foregroundColor: string): void {
    this.setCellsStyle(sheet, area, () => ({ fill: { foregroundColor } }));
  }

  setTextColor(sheet: number, area: Area, color: string): void {
    this.setCellsStyle(sheet, area, () => ({ font: { color } }));
  }

  toggleAlign(
    sheet: number,
    area: Area,
    alignment: NonNullable<
      NonNullable<Parameters<ICell['style']['bulkUpdate']>[0]['alignment']>['horizontalAlignment']
    >,
  ): void {
    this.setCellsStyle(sheet, area, (style) => ({
      alignment: {
        horizontalAlignment:
          style.alignment.horizontalAlignment === alignment ? 'general' : alignment,
      },
    }));
  }

  toggleFontStyle(sheet: number, area: Area, fontStyle: keyof ICell['style']['font']): void {
    this.setCellsStyle(sheet, area, (style) => ({
      font: {
        [fontStyle]: !style.font[fontStyle],
      },
    }));
  }

  isQuotePrefix(sheet: number, row: number, column: number): boolean {
    const style = this.getCellStyle(sheet, row, column);
    // FIXME
    return false;
  }

  getFormulaOrValue(sheet: number, row: number, column: number): string {
    const cell = this.workbook.cell(sheet, row, column);
    return cell.formula ?? `${cell.value ?? ''}`;
  }

  hasFormula(sheet: number, row: number, column: number): boolean {
    return this.workbook.cell(sheet, row, column).formula !== null;
  }

  canUndo(): boolean {
    return this.history.undo.length > 0;
  }

  undo(): void {
    /* const diffs = this.history.undo.pop();
    if (!diffs) {
      return;
    }
    const { workbook } = this;
    this.history.redo.push(diffs);
    let forceEvaluate = false;
    for (const diff of diffs) {
      switch (diff.type) {
        case 'set_cell_value': {
          if (diff.oldValue !== null) {
            wasm.set_input(diff.sheet, diff.row, diff.column, diff.oldValue, diff.oldStyle);
          } else {
            wasm.set_input(diff.sheet, diff.row, diff.column, '', diff.oldStyle);
            wasm.delete_cell(diff.sheet, diff.row, diff.column);
          }
          forceEvaluate = true;
          break;
        }
        case 'set_column_width': {
          wasm.set_column_width(diff.sheet, diff.column, diff.oldValue);
          break;
        }
        case 'set_row_height': {
          wasm.set_row_height(diff.sheet, diff.row, diff.oldValue);
          break;
        }
        case 'delete_cell': {
          if (diff.oldValue !== null) {
            wasm.set_input(diff.sheet, diff.row, diff.column, diff.oldValue, diff.oldStyle);
            forceEvaluate = true;
          }
          break;
        }
        case 'remove_cell': {
          if (diff.oldValue !== null) {
            wasm.set_input(diff.sheet, diff.row, diff.column, diff.oldValue, diff.oldStyle);
            forceEvaluate = true;
          } else {
            wasm.set_input(diff.sheet, diff.row, diff.column, '', diff.oldStyle);
            wasm.delete_cell(diff.sheet, diff.row, diff.column);
          }
          break;
        }
        case 'set_cell_style': {
          wasm.set_cell_style(diff.sheet, diff.row, diff.column, diff.oldValue);
          break;
        }
        case 'insert_row': {
          wasm.delete_rows(diff.sheet, diff.row, 1);
          forceEvaluate = true;
          break;
        }
        case 'delete_row': {
          wasm.insert_rows(diff.sheet, diff.row, 1);
          const { rowData, rowHeight } = diff.oldValue;
          wasm.set_row_height(diff.sheet, diff.row, rowHeight);
          wasm.set_row_undo_data(diff.sheet, diff.row, rowData);
          forceEvaluate = true;
          break;
        }
        /* istanbul ignore next */
    /* default: {
          const unrecognized: never = diff;
          throw new Error(`Unrecognized diff type - ${unrecognized}}.`);
        }
      }
    }
    if (forceEvaluate) {
      wasm.evaluate();
    } */
  }

  canRedo(): boolean {
    return this.history.redo.length > 0;
  }

  redo(): void {
    /* const diffs = this.history.redo.pop();
    if (!diffs) {
      return;
    }
    const { wasm } = this;
    this.history.undo.push(diffs);
    let forceEvaluate = false;
    for (const diff of diffs) {
      switch (diff.type) {
        case 'set_cell_value': {
          if (diff.newValue !== null) {
            wasm.set_input(diff.sheet, diff.row, diff.column, diff.newValue, diff.newStyle);
          } else {
            // we need to set the style
            wasm.set_input(diff.sheet, diff.row, diff.column, '', diff.newStyle);
            wasm.delete_cell(diff.sheet, diff.row, diff.column);
          }
          forceEvaluate = true;
          break;
        }
        case 'set_column_width': {
          wasm.set_column_width(diff.sheet, diff.column, diff.newValue);
          break;
        }
        case 'set_row_height': {
          wasm.set_row_height(diff.sheet, diff.row, diff.newValue);
          break;
        }
        case 'delete_cell': {
          wasm.delete_cell(diff.sheet, diff.row, diff.column);
          forceEvaluate = true;
          break;
        }
        case 'remove_cell': {
          wasm.remove_cell(diff.sheet, diff.row, diff.column);
          forceEvaluate = true;
          break;
        }
        case 'set_cell_style': {
          wasm.set_cell_style(diff.sheet, diff.row, diff.column, diff.newValue);
          break;
        }
        case 'insert_row': {
          wasm.insert_rows(diff.sheet, diff.row, 1);
          forceEvaluate = true;
          break;
        }
        case 'delete_row': {
          wasm.delete_rows(diff.sheet, diff.row, 1);
          forceEvaluate = true;
          break;
        }
        /* istanbul ignore next */
    /* default: {
          const unrecognized: never = diff;
          throw new Error(`Unrecognized diff type - ${unrecognized}}.`);
        }
      }
    }
    if (forceEvaluate) {
      wasm.evaluate();
    } */
  }

  deleteCells(sheet: number, area: Area): void {
    const diffs: Diff[] = [];
    for (let row = area.rowStart; row <= area.rowEnd; row += 1) {
      for (let column = area.columnStart; column <= area.columnEnd; column += 1) {
        // TODO: A lot of evaluations
        this.workbook.cell(sheet, row, column).value = null;
      }
    }
    this.history.undo.push(diffs);
    this.history.redo = [];
    this.notifySubscribers({ type: 'deleteCells' });
  }

  setCellValue(sheet: number, row: number, column: number, value: string): void {
    this.workbook.cell(sheet, row, column).input = value;
    this.notifySubscribers({ type: 'setCellValue' });
  }

  extendTo(sheet: number, sourceArea: Area, targetArea: Area): void {
    const sourceRowBoundaries = [sourceArea.rowStart, sourceArea.rowEnd];
    const sourceColumnBoundaries = [sourceArea.columnStart, sourceArea.columnEnd];

    const [sourceRowStart, sourceRowEnd] = [
      Math.min(...sourceRowBoundaries),
      Math.max(...sourceRowBoundaries),
    ];
    const [sourceColumnStart, sourceColumnEnd] = [
      Math.min(...sourceColumnBoundaries),
      Math.max(...sourceColumnBoundaries),
    ];

    const shouldReverseRows = sourceRowStart > targetArea.rowStart;
    const shouldReverseColumns = sourceColumnStart > targetArea.columnStart;

    for (let rowOffset = 0; rowOffset <= targetArea.rowEnd - targetArea.rowStart; rowOffset += 1) {
      for (
        let columnOffset = 0;
        columnOffset <= targetArea.columnEnd - targetArea.columnStart;
        columnOffset += 1
      ) {
        const sourceRowOffset = rowOffset % (sourceRowEnd - sourceRowStart + 1);
        const sourceColumnOffset = columnOffset % (sourceColumnEnd - sourceColumnStart + 1);

        const sourceRow = !shouldReverseRows
          ? sourceRowStart + sourceRowOffset
          : sourceRowEnd - sourceRowOffset;

        const sourceColumn = !shouldReverseColumns
          ? sourceColumnStart + sourceColumnOffset
          : sourceColumnEnd - sourceColumnOffset;

        const targetRow = !shouldReverseRows
          ? targetArea.rowStart + rowOffset
          : targetArea.rowEnd - rowOffset;

        const targetColumn = !shouldReverseColumns
          ? targetArea.columnStart + columnOffset
          : targetArea.columnEnd - columnOffset;

        this.extendToCell(sheet, sourceRow, sourceColumn, targetRow, targetColumn);
      }
    }
    this.notifySubscribers({ type: 'extendTo' });
  }

  private extendToCell(
    sheet: number,
    sourceRow: number,
    sourceColumn: number,
    targetRow: number,
    targetColumn: number,
  ) {
    const extendedValue = this.workbook.sheets
      .get(sheet)
      .userInterface.getExtendedValue(sourceRow, sourceColumn, targetRow, targetColumn);
    this.workbook.cell(sheet, targetRow, targetColumn).input = extendedValue;
    const style = this.getCellStyle(sheet, sourceRow, sourceColumn);
    // TODO: Probably copying style should be supported in the SDK
    this.setCellsStyle(
      sheet,
      {
        rowStart: targetRow,
        rowEnd: targetRow,
        columnStart: targetColumn,
        columnEnd: targetColumn,
      },
      () => ({
        font: {
          bold: style.font.bold,
          italics: style.font.italics,
          color: style.font.color,
          underline: style.font.underline,
          strikethrough: style.font.strikethrough,
        },
        fill: {
          foregroundColor: style.fill.foregroundColor,
          backgroundColor: style.fill.backgroundColor,
          patternType: style.fill.patternType,
        },
        numberFormat: style.numberFormat,
        alignment: {
          horizontalAlignment: style.alignment.horizontalAlignment,
          verticalAlignment: style.alignment.verticalAlignment,
          wrapText: style.alignment.wrapText,
        },
      }),
    );
  }

  getFormulaValueOrNull(sheet: number, row: number, column: number): string | null {
    // FIXME: We need cell.input getter
    const cell = this.workbook.cell(sheet, row, column);
    return cell.formula ?? String(cell.value ?? '');
  }

  getFrozenRowsCount(sheet: number): number {
    return 0;
    /*
    const result = this.wasm.get_frozen_rows(sheet);
    if (!result.success) {
      return 0;
    }
    return result.value;
    */
  }

  getFrozenColumnsCount(sheet: number): number {
    return 0;
    /*
    const result = this.wasm.get_frozen_columns(sheet);
    if (!result.success) {
      return 0;
    }
    return result.value;
    */
  }

  copy(area: SheetArea): { tsv: string; area: SheetArea; sheetData: SheetData } {
    const { sheet, rowStart, columnStart, rowEnd, columnEnd } = area;
    const sheetName = this.workbook.sheets.get(sheet).name;

    const tsv = [];
    const sheetData: SheetData = { data: {}, sheetName };

    for (let row = rowStart; row <= rowEnd; row += 1) {
      const tsvRow = [];
      sheetData.data[row] = {};

      for (let column = columnStart; column <= columnEnd; column += 1) {
        const value = this.getFormulaValueOrNull(sheet, row, column); // FIXME
        const text = this.getTextAt(sheet, row, column);
        const style = null; // this.getCellStyle(sheet, row, column); // TODO: stringify

        sheetData.data[row][column] = { value, style };
        tsvRow.push(escapeTabs(text));
      }

      tsv.push(tsvRow.join('\t'));
    }
    return { tsv: tsv.join('\n'), area, sheetData };
  }

  // This takes care of _internal_ paste, that is copy/paste within the application
  paste(source: SheetArea, target: SheetArea, sheetData: SheetData, type: PasteType): void {
    const sourceArea = {
      sheet: source.sheet,
      row: source.rowStart,
      column: source.columnStart,
      width: source.columnEnd - source.columnStart + 1,
      height: source.rowEnd - source.rowStart + 1,
    };

    const deltaRow = target.rowStart - source.rowStart;
    const deltaColumn = target.columnStart - source.columnStart;

    for (let row = source.rowStart; row <= source.rowEnd; row += 1) {
      for (let column = source.columnStart; column <= source.columnEnd; column += 1) {
        const cellData = getValueFromSheetData(sheetData, row, column);
        const targetRow = row + deltaRow;
        const targetColumn = column + deltaColumn;

        const cell = this.workbook.cell(target.sheet, targetRow, targetColumn);
        // cell.style = cellData.style;

        if (cellData.value === null) {
          cell.input = '';
        } else {
          switch (type) {
            case 'copy': {
              cell.input = this.workbook.getCopiedValueExtended(
                cellData.value,
                sheetData.sheetName,
                { sheet: source.sheet, row, column },
                { sheet: target.sheet, row: targetRow, column: targetColumn },
              );
              break;
            }
            case 'cut': {
              cell.input = this.workbook.getCutValueMoved(
                cellData.value,
                { sheet: source.sheet, row, column },
                { sheet: target.sheet, row: targetRow, column: targetColumn },
                sourceArea,
              );
              break;
            }
            /* istanbul ignore next */
            default: {
              const unrecognized: never = type;
              throw new Error(`Unrecognized paste type - ${unrecognized}}.`);
            }
          }
        }
      }
    }

    if (type === 'cut') {
      const targetArea = {
        sheet: target.sheet,
        rowStart: target.rowStart,
        rowEnd: target.rowStart + source.rowEnd - source.rowStart,
        columnStart: target.columnStart,
        columnEnd: target.columnStart + source.columnEnd - source.columnStart,
      };
      // remove all cells that are not in the paste area
      for (let row = source.rowStart; row <= source.rowEnd; row += 1) {
        for (let column = source.columnStart; column <= source.columnEnd; column += 1) {
          if (!isCellInArea(source.sheet, row, column, targetArea)) {
            this.workbook.cell(source.sheet, row, column).delete();
          }
        }
      }
      // Change all formulas that pointed to the old area to the new area
      this.workbook.forwardReferences(sourceArea, {
        sheet: target.sheet,
        row: target.rowStart,
        column: target.columnStart,
      });
    }
    this.notifySubscribers({ type: 'paste' });
  }

  pasteText(sheet: number, target: Cell, value: string): void {
    const parsedData = Papa.parse(value, { delimiter: '\t', header: false });
    const data = parsedData.data as string[][];

    for (let deltaRow = 0; deltaRow < data.length; deltaRow += 1) {
      const line = data[deltaRow];
      const targetRow = target.row + deltaRow;

      for (let deltaColumn = 0; deltaColumn < line.length; deltaColumn += 1) {
        const newValue = line[deltaColumn].trim();
        const targetColumn = target.column + deltaColumn;
        this.workbook.cell(sheet, targetRow, targetColumn).input = newValue;
      }
    }
    this.notifySubscribers({ type: 'pasteText' });
  }

  getNavigationEdge(
    key: NavigationKey,
    sheet: number,
    row: number,
    column: number,
  ): { row: number; column: number } {
    const keyToDirection = {
      ArrowUp: 'up',
      ArrowDown: 'down',
      ArrowLeft: 'left',
      ArrowRight: 'right',
      // FIXME
      Home: 'up',
      End: 'up',
    } satisfies Record<NavigationKey, NavigationDirection>;
    const [newRow, newColumn] = this.workbook.sheets
      .get(sheet)
      .userInterface.navigateToEdgeInDirection(row, column, keyToDirection[key]);
    return {
      row: Math.min(newRow, workbookLastRow),
      column: Math.min(newColumn, workbookLastColumn),
    };
  }

  insertRow(sheet: number, row: number): void {
    this.workbook.sheets.get(sheet).insertRows(row, 1);
    this.notifySubscribers({ type: 'insertRow' });
  }

  // We need to delete (and save) the row style
  // We need to delete (and save) the data
  deleteRow(sheet: number, row: number): void {
    this.workbook.sheets.get(sheet).deleteRows(row, 1);
    this.notifySubscribers({ type: 'deleteRow' });
  }

  insertColumn(sheet: number, column: number): void {
    this.workbook.sheets.get(sheet).insertColumns(column, 1);
    this.notifySubscribers({ type: 'insertRow' });
  }

  // We need to delete (and save) the column style
  // We need to delete (and save) the data
  deleteColumn(sheet: number, column: number): void {
    this.workbook.sheets.get(sheet).deleteColumns(column, 1);
    this.notifySubscribers({ type: 'deleteRow' });
  }

  saveToXlsx(): Uint8Array {
    return this.workbook.saveToXlsx();
  }
}
