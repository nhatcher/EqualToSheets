import { ICell, ISheet } from '@equalto-software/calc';

export interface Area {
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
}

interface AreaWithBorderInterface extends Area {
  border: 'left' | 'top' | 'right' | 'bottom';
}

export type AreaWithBorder = AreaWithBorderInterface | null;

export type StyleReducer = (style: ICell['style']) => Parameters<ICell['style']['bulkUpdate']>[0];

export interface IModel {
  setCellValue(sheet: number, row: number, column: number, value: string): void;
  setCellsStyle(sheet: number, area: Area, reducer: StyleReducer): void;
  insertRow(sheet: number, row: number): void;
  setRowHeight(sheet: number, row: number, height: number): void;
  deleteRow(sheet: number, row: number): void;
  insertColumn(sheet: number, column: number): void;
  setColumnWidth(sheet: number, column: number, width: number): void;
  deleteColumn(sheet: number, column: number): void;
  addBlankSheet(): ISheet;
  renameSheet(sheet: number, newName: string): void;
  deleteSheet(sheet: number): void;
  deleteCells(sheet: number, area: Area): void;
  enableHistory(): void;
  disableHistory(): void;
}
