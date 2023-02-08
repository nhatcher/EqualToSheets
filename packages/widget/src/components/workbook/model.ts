/* eslint-disable class-methods-use-this */
/* eslint-disable @typescript-eslint/no-unused-vars */

import { ICell, IWorkbook, NavigationDirection } from '@equalto-software/calc';
import Papa from 'papaparse';
import { TabsInput } from './components/navigation/common';
import { workbookLastColumn, workbookLastRow } from './constants';
import { MarkedToken } from './tokenTypes';
import { Area, Cell, NavigationKey } from './util';

export enum ValueType {
  Boolean = 'boolean',
  Number = 'number',
  Text = 'text',
}

export type ExcelValue = string | number | boolean;

interface ModelSettings {
  workbook: IWorkbook;
  readOnly?: boolean;
  getTokens: (formula: string) => MarkedToken[];
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

interface Color {
  RGB: string;
}

interface CellStyleFill {
  pattern_type: string;
  fg_color?: Color;
  bg_color?: Color;
}

interface CellStyleFont {
  u: boolean;
  b: boolean;
  i: boolean;
  strike: boolean;
  sz: number;
  color: Color;
  name: string;
  family: number;
  scheme: string;
}

interface BorderItem {
  style: string;
  color: Color;
}
interface CellStyleBorder {
  diagonal_up?: boolean;
  diagonal_down?: boolean;
  left: BorderItem;
  right: BorderItem;
  top: BorderItem;
  bottom: BorderItem;
  diagonal: BorderItem;
}
export interface CellStyle {
  read_only: boolean;
  quote_prefix: boolean;
  fill: CellStyleFill;
  font: CellStyleFont;
  border: CellStyleBorder;
  num_fmt: string;
  horizontal_alignment: string;
}

interface CellData {
  value: CellValue;
  style: CellStyle | null;
}

interface RowData {
  [column: number]: CellData;
}

interface SheetData {
  [row: number]: RowData;
}

type StyleReducer = (style: CellStyle) => CellStyle;

type PasteType = 'copy' | 'cut';

type FontStyle = 'underline' | 'italic' | 'bold' | 'strikethrough';

// Replaces all tabs with spaces in a string
function escapeTabs(s: string): string {
  return s.replace(/\t/g, '  ');
}

function getValueFromSheetData(sheetData: SheetData, row: number, column: number): CellData {
  const rowData = sheetData[row];
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
  getTokens: (formula: string) => MarkedToken[];

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
    // TODO: return this.wasm.get_column_width(sheet, column);
    return 100;
  }

  setColumnWidth(sheet: number, column: number, width: number): void {
    this.history.undo.push([
      {
        type: 'set_column_width',
        sheet,
        column,
        newValue: width,
        oldValue: 100, // TODO: this.wasm.get_column_width(sheet, column),
      },
    ]);
    this.history.redo = [];
    // TODO:
    // this.wasm.set_column_width(sheet, column, width);
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

  getRowHeight(sheet: number, rowIndex: number): number {
    return 20;
    // TODO:
    // return this.wasm.get_row_height(sheet, rowIndex);
  }

  setRowHeight(sheet: number, row: number, height: number): void {
    this.history.undo.push([
      {
        type: 'set_row_height',
        sheet,
        row,
        newValue: height,
        oldValue: 20, // TODO: this.wasm.get_row_height(sheet, row),
      },
    ]);
    this.history.redo = [];
    // this.wasm.set_row_height(sheet, row, height);
    this.notifySubscribers({ type: 'setRowHeight' });
  }

  getTextAt(sheet: number, row: number, column: number): string {
    const cell = this.workbook.cell(sheet, row, column);
    return cell.formula ?? `${cell.value ?? ''}`;
  }

  getUICell(sheet: number, row: number, column: number): ICell {
    return this.workbook.cell(sheet, row, column);
  }

  getCellStyle(sheet: number, row: number, column: number): CellStyle {
    return {
      fill: { fg_color: { RGB: '#FFFFFF' }, pattern_type: 'solid' },
      font: {
        color: { RGB: '#000000' },
        b: false,
        i: false,
        strike: false,
        u: false,
        name: 'Arial',
        family: 0,
        sz: 14,
        scheme: 'minor',
      },
      horizontal_alignment: 'left',
      num_fmt: 'general',
      read_only: false,
      quote_prefix: false,
      border: {
        left: {
          style: 'solid',
          color: { RGB: '#FFFFFF' },
        },
        right: {
          style: 'solid',
          color: { RGB: '#FFFFFF' },
        },
        top: {
          style: 'solid',
          color: { RGB: '#FFFFFF' },
        },
        bottom: {
          style: 'solid',
          color: { RGB: '#FFFFFF' },
        },
        diagonal: {
          style: 'none',
          color: { RGB: '#FFFFFF' },
        },
      },
    };
    // TODO:
    // return JSON.parse(this.wasm.get_style_for_cell(sheet, row, column));
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
    this.notifySubscribers({ type: 'setCellStyle' });
    // TODO:
    /*  const diffs: Diff[] = [];
    if (this.isAreaReadOnly(sheet, area)) {
      return;
    }
    let { rowStart, rowEnd, columnStart, columnEnd } = area;
    if (rowStart > rowEnd) {
      [rowStart, rowEnd] = [rowEnd, rowStart];
    }
    if (columnStart > columnEnd) {
      [columnStart, columnEnd] = [columnEnd, columnStart];
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const styleString = this.getCellStyle(sheet, row, column) ?? 0;
        const style = JSON.parse(styleString);
        const newStyle = reducer(style);
        diffs.push({
          type: 'set_cell_style',
          sheet,
          row,
          column,
          newValue: newStyle,
          oldValue: style,
        });
        this.wasm.set_cell_style(sheet, row, column, newStyle);
      }
    }
    this.history.undo.push(diffs);
    this.history.redo = []; */
  }

  setNumberFormat(sheet: number, area: Area, numberFormat: string): void {
    this.setCellsStyle(sheet, area, (style) => ({
      ...style,
      num_fmt: numberFormat,
    }));
  }

  setFillColor(sheet: number, area: Area, color: string): void {
    this.setCellsStyle(sheet, area, (style) => ({
      ...style,
      fill: {
        ...style.fill,
        fg_color: {
          RGB: color,
        },
      },
    }));
  }

  setTextColor(sheet: number, area: Area, color: string): void {
    this.setCellsStyle(sheet, area, (style) => ({
      ...style,
      font: {
        ...style.font,
        color: {
          RGB: color,
        },
      },
    }));
  }

  toggleAlign(sheet: number, area: Area, alignment: 'left' | 'center' | 'right'): void {
    this.setCellsStyle(sheet, area, (style) => ({
      ...style,
      horizontal_alignment: style.horizontal_alignment === alignment ? 'default' : alignment,
    }));
  }

  toggleFontStyle(sheet: number, area: Area, fontStyle: FontStyle): void {
    const propertyMap: Record<FontStyle, keyof CellStyleFont> = {
      underline: 'u',
      italic: 'i',
      bold: 'b',
      strikethrough: 'strike',
    };
    this.setCellsStyle(sheet, area, (style) => ({
      ...style,
      font: {
        ...style.font,
        [propertyMap[fontStyle]]: !style.font[propertyMap[fontStyle]],
      },
    }));
  }

  isCellReadOnly(sheet: number, row: number, column: number): boolean {
    return false;
    // TODO:
    // const style = JSON.parse(this.wasm.get_style_for_cell(sheet, row, column));
    // return style.read_only || this.readOnly;
  }

  isQuotePrefix(sheet: number, row: number, column: number): boolean {
    const style = this.getCellStyle(sheet, row, column);
    return style.quote_prefix;
  }

  isAreaReadOnly(sheet: number, area: Area): boolean {
    for (let row = area.rowStart; row <= area.rowEnd; row += 1) {
      for (let column = area.columnStart; column <= area.columnEnd; column += 1) {
        if (this.isCellReadOnly(sheet, row, column)) {
          return true;
        }
      }
    }
    return false;
  }

  isRowReadOnly(sheet: number, row: number): boolean {
    return false;
    // TODO:
    // return this.wasm.is_row_read_only(sheet, row);
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
    if (this.isAreaReadOnly(sheet, area)) {
      return;
    }
    for (let row = area.rowStart; row <= area.rowEnd; row += 1) {
      for (let column = area.columnStart; column <= area.columnEnd; column += 1) {
        /* const oldValue = this.getFormulaValueOrNull(sheet, row, column);
        diffs.push({
          type: 'delete_cell',
          sheet,
          row,
          column,
          oldValue,
          oldStyle: .get_cell_style_index(sheet, row, column),
        }); */
        // TODO: A lot of evaluations
        this.workbook.cell(sheet, row, column).value = null;
      }
    }
    this.history.undo.push(diffs);
    this.history.redo = [];
    this.notifySubscribers({ type: 'deleteCells' });
  }

  setCellValue(sheet: number, row: number, column: number, value: string): void {
    if (this.isCellReadOnly(sheet, row, column)) {
      return;
    }
    const styleIndex = 0; // TODO: this.wasm.get_cell_style_index(sheet, row, column);
    this.history.undo.push([
      {
        type: 'set_cell_value',
        sheet,
        row,
        column,
        newStyle: styleIndex,
        newValue: value,
        oldValue: this.getFormulaValueOrNull(sheet, row, column),
        oldStyle: styleIndex,
      },
    ]);
    this.history.redo = [];
    if (value.startsWith('=')) {
      this.workbook.cell(sheet, row, column).formula = value;
    } else {
      this.workbook.cell(sheet, row, column).value = value;
    }
    this.notifySubscribers({ type: 'setCellValue' });
  }

  extendTo(sheet: number, initialArea: Area, extendedArea: Area): void {
    this.notifySubscribers({ type: 'extendTo' });
    /* const diffs: Array<Diff> = [];
    let { rowStart, rowEnd, columnStart, columnEnd } = initialArea;
    if (rowStart > rowEnd) {
      [rowStart, rowEnd] = [rowEnd, rowStart];
    }
    if (columnStart > columnEnd) {
      [columnStart, columnEnd] = [columnEnd, columnStart];
    }
    if (columnStart === extendedArea.columnStart && columnEnd === extendedArea.columnEnd) {
      // extend by rows
      let offsetRow;
      let startRow;
      if (rowEnd + 1 === extendedArea.rowStart) {
        offsetRow = extendedArea.rowStart;
        startRow = rowStart;
      } else {
        offsetRow = extendedArea.rowEnd;
        startRow = rowEnd;
      }

      for (let row = extendedArea.rowStart; row <= extendedArea.rowEnd; row += 1) {
        for (let column = columnStart; column <= columnEnd; column += 1) {
          if (!this.isCellReadOnly(sheet, row, column)) {
            const sourceRow = startRow + ((row - offsetRow) % (rowEnd - rowStart + 1));
            let value = this.wasm.extend_to(sheet, sourceRow, column, row, column);
            // We don't copy over read-only styles
            const oldStyle = this.wasm.get_cell_style_index(sheet, row, column);
            const newStyle = this.isCellReadOnly(sheet, sourceRow, column)
              ? oldStyle
              : this.wasm.get_cell_style_index(sheet, sourceRow, column);
            if (this.isQuotePrefix(sheet, sourceRow, column) && !value.startsWith("'")) {
              value = `'${value}`;
            }
            diffs.push({
              type: 'set_cell_value',
              sheet,
              row,
              column,
              newValue: value,
              newStyle,
              oldValue: this.getFormulaValueOrNull(sheet, row, column),
              oldStyle,
            });
            this.wasm.set_input(sheet, row, column, value, newStyle);
          }
        }
      }
    } else if (rowStart === extendedArea.rowStart && rowEnd === extendedArea.rowEnd) {
      // extend by columns
      let offsetColumn;
      let startColumn;
      if (columnEnd + 1 === extendedArea.columnStart) {
        offsetColumn = extendedArea.columnStart;
        startColumn = columnStart;
      } else {
        offsetColumn = extendedArea.columnEnd;
        startColumn = columnEnd;
      }
      for (let row = rowStart; row <= rowEnd; row += 1) {
        for (let column = extendedArea.columnStart; column <= extendedArea.columnEnd; column += 1) {
          if (!this.isCellReadOnly(sheet, row, column)) {
            const sourceColumn =
              startColumn + ((column - offsetColumn) % (columnEnd - columnStart + 1));

            // We don't copy over read-only styles
            const oldStyle = this.wasm.get_cell_style_index(sheet, row, column);
            const newStyle = this.isCellReadOnly(sheet, row, sourceColumn)
              ? oldStyle
              : this.wasm.get_cell_style_index(sheet, row, sourceColumn);

            let value = this.wasm.extend_to(sheet, row, sourceColumn, row, column);
            if (this.isQuotePrefix(sheet, row, sourceColumn) && !value.startsWith("'")) {
              value = `'${value}`;
            }
            diffs.push({
              type: 'set_cell_value',
              sheet,
              row,
              column,
              newValue: value,
              oldValue: this.getFormulaValueOrNull(sheet, row, column),
              newStyle,
              oldStyle,
            });
            this.wasm.set_input(sheet, row, column, value, newStyle);
          }
        }
      }
    }
    this.history.undo.push(diffs);
    this.history.redo = [];
    this.wasm.evaluate(); */
  }

  copy(area: SheetArea): { tsv: string; area: SheetArea; sheetData: SheetData } {
    const { sheet, rowStart, columnStart, rowEnd, columnEnd } = area;
    const tsv = [];
    const sheetData: SheetData = {};
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const tsvRow = [];
      sheetData[row] = {};
      for (let column = columnStart; column <= columnEnd; column += 1) {
        // If we would want to copy the formulas instead we would use this line:
        const value = this.getFormulaValueOrNull(sheet, row, column);
        const text = this.getTextAt(sheet, row, column);
        const style = this.getCellStyle(sheet, row, column);
        sheetData[row][column] = { value, style };
        tsvRow.push(escapeTabs(text));
      }
      tsv.push(tsvRow.join('\t'));
    }
    return { tsv: tsv.join('\n'), area, sheetData };
  }

  getFormulaValueOrNull(sheet: number, row: number, column: number): CellValue {
    const cell = this.workbook.cell(sheet, row, column);
    if (typeof cell.value === null) {
      return null;
    }
    return cell.formula ?? cell.value;
  }

  getStyleIndexOrCreate(style: CellStyle | null): number {
    return 0;
    /*
    if (style === null) {
      return 0;
    }
    const result = this.wasm.get_style_index_or_create(style);
    if (!result.success) {
      return 0;
    }
    return result.index;
    */
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

  // This takes care of _internal_ paste, that is copy/paste within the application
  paste(source: SheetArea, target: SheetArea, sheetData: SheetData, type: PasteType): void {
    this.notifySubscribers({ type: 'paste' });
    /* const sourceArea = {
      sheet: source.sheet,
      row: source.rowStart,
      column: source.columnStart,
      width: source.columnEnd - source.columnStart + 1,
      height: source.rowEnd - source.rowStart + 1,
    };

    const deltaRow = target.rowStart - source.rowStart;
    const deltaColumn = target.columnStart - source.columnStart;
    let diffs: Diff[] = [];
    for (let row = source.rowStart; row <= source.rowEnd; row += 1) {
      for (let column = source.columnStart; column <= source.columnEnd; column += 1) {
        const targetRow = row + deltaRow;
        const targetColumn = column + deltaColumn;
        if (!this.isCellReadOnly(target.sheet, targetRow, targetColumn)) {
          const cellData = getValueFromSheetData(sheetData, row, column);
          // We don't copy over read-only styles
          const oldStyle = 0;
          // TODO: const oldStyle = this.wasm.get_cell_style_index(target.sheet, targetRow, targetColumn);
          const newStyle = this.isCellReadOnly(target.sheet, row, column)
            ? oldStyle
            : this.getStyleIndexOrCreate(cellData.style);
          if (cellData.value === null) {
            // HACK. We are pasting an empty cell. This means we need to delete the content and update the style
            // At the time of writing we don't have an API for that so:
            // i.  We set some value in the cell to update the style
            // ii. We delete the content
            const oldValue = this.getFormulaValueOrNull(target.sheet, targetRow, targetColumn);
            diffs.push(
              {
                type: 'set_cell_value',
                sheet: target.sheet,
                row: targetRow,
                column: targetColumn,
                newValue: '',
                oldValue,
                newStyle,
                oldStyle,
              },
              {
                type: 'delete_cell',
                sheet: target.sheet,
                row: targetRow,
                column: targetColumn,
                oldValue,
                oldStyle,
              },
            );
            this.workbook.cell(target.sheet, targetRow, targetColumn).value = '';
            this.wasm.set_input(target.sheet, targetRow, targetColumn, '', newStyle);
            this.wasm.delete_cell(target.sheet, targetRow, targetColumn);
          } else {
            let value;
            switch (type) {
              case 'copy': {
                const result = JSON.parse(
                  this.wasm.extend_formula_to(
                    JSON.stringify({
                      source: {
                        sheet: source.sheet,
                        row,
                        column,
                      },
                      value: cellData.value,
                      target: {
                        sheet: target.sheet,
                        row: targetRow,
                        column: targetColumn,
                      },
                    }),
                  ),
                );
                value = result.value;
                break;
              }
              case 'cut': {
                const result = JSON.parse(
                  this.wasm.move_cell_value_to_area(
                    JSON.stringify({
                      source: {
                        sheet: source.sheet,
                        row,
                        column,
                      },
                      value: cellData.value,
                      target: {
                        sheet: target.sheet,
                        row: targetRow,
                        column: targetColumn,
                      },
                      area: sourceArea,
                    }),
                  ),
                );
                value = result.value;
                break;
              }
              /* istanbul ignore next */
    /* default: {
                const unrecognized: never = type;
                throw new Error(`Unrecognized paste type - ${unrecognized}}.`);
              }
            }
            if (
              this.isQuotePrefix(source.sheet, row, column) &&
              !(value as string).startsWith("'")
            ) {
              value = `'${value}`;
            }
            diffs.push({
              type: 'set_cell_value',
              sheet: target.sheet,
              row: targetRow,
              column: targetColumn,
              newValue: value,
              oldValue: this.getFormulaValueOrNull(target.sheet, targetRow, targetColumn),
              newStyle,
              oldStyle,
            });
            this.wasm.set_input(target.sheet, targetRow, targetColumn, value, newStyle);
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
          if (
            !this.isCellReadOnly(source.sheet, row, column) &&
            !isCellInArea(source.sheet, row, column, targetArea)
          ) {
            diffs.push({
              type: 'remove_cell',
              sheet: source.sheet,
              row,
              column,
              oldValue: this.getFormulaValueOrNull(source.sheet, row, column),
              oldStyle: this.wasm.get_cell_style_index(source.sheet, row, column),
            });
            this.wasm.remove_cell(source.sheet, row, column);
          }
        }
      }
      // Change all formulas that pointed to the old area to the new area
      const changes = JSON.parse(
        this.wasm.forward_references(
          JSON.stringify(sourceArea),
          target.sheet,
          target.rowStart,
          target.columnStart,
        ),
      );
      if (changes.success) {
        diffs = [...diffs, ...changes.diffList];
      }
    }
    this.history.undo.push(diffs);
    this.history.redo = [];
    this.wasm.evaluate();  */
  }

  pasteText(sheet: number, target: Cell, value: string): void {
    this.notifySubscribers({ type: 'pasteText' });
    /* const parsedData = Papa.parse(value, { delimiter: '\t', header: false });
    const data = parsedData.data as string[][];
    const rowCount = data.length;
    const diffs: Diff[] = [];
    for (let deltaRow = 0; deltaRow < rowCount; deltaRow += 1) {
      const line = data[deltaRow];
      const columnCount = line.length;
      const targetRow = target.row + deltaRow;
      for (let deltaColumn = 0; deltaColumn < columnCount; deltaColumn += 1) {
        const newValue = line[deltaColumn].trim();
        const targetColumn = target.column + deltaColumn;
        if (!this.isCellReadOnly(sheet, targetRow, targetColumn)) {
          const styleIndex = this.wasm.get_cell_style_index(sheet, targetRow, targetColumn);
          diffs.push({
            type: 'set_cell_value',
            sheet,
            row: targetRow,
            column: targetColumn,
            newValue,
            oldValue: this.getFormulaValueOrNull(sheet, targetRow, targetColumn),
            newStyle: styleIndex,
            oldStyle: styleIndex,
          });
          this.wasm.set_input(sheet, targetRow, targetColumn, newValue, styleIndex);
        }
      }
    }
    this.history.undo.push(diffs);
    this.history.redo = [];
    this.wasm.evaluate(); */
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
    this.notifySubscribers({ type: 'insertRow' });
    /* this.wasm.insert_rows(sheet, row, 1);
    this.history.undo.push([
      {
        type: 'insert_row',
        sheet,
        row,
      },
    ]);
    this.history.redo = [];
    this.wasm.evaluate(); */
  }

  // We need to delete (and save) the row style
  // We need to delete (and save) the data
  deleteRow(sheet: number, row: number): void {
    this.notifySubscribers({ type: 'deleteRow' });
    // get old row style, if any
    // oldRowStyle
    /* const rowData = this.wasm.get_row_undo_data(sheet, row);
    const rowHeight = this.wasm.get_row_height(sheet, row);
    this.wasm.delete_rows(sheet, row, 1);
    this.history.undo.push([
      {
        type: 'delete_row',
        sheet,
        row,
        oldValue: { rowHeight, rowData },
      },
    ]);
    this.history.redo = [];
    this.wasm.evaluate(); */
  }
}
