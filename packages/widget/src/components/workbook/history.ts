import { CellStyleSnapshot } from '@equalto-software/calc';
import { IModel } from './types';

// TODO: Move to sheet ids or just use indexes for now

type DeletedCell = {
  cell: Cell;
  value: string | null;
  style: CellStyleSnapshot;
};

export class ActionHistory {
  private enabled: boolean;

  private undoStack: IAction[][];

  private redoStack: IAction[][];

  constructor() {
    this.undoStack = [];
    this.redoStack = [];
    this.enabled = true;
  }

  enable() {
    this.enabled = true;
  }

  disable() {
    this.enabled = false;
  }

  push(actions: IAction[]) {
    if (!this.enabled) {
      return;
    }
    this.undoStack.push(actions);
    this.redoStack = [];
  }

  undo() {
    const actions = this.undoStack.pop();
    if (actions) {
      const redoActions: IAction[] = [];
      actions.forEach((action) => {
        action.undo();
        redoActions.push(action);
      });
      this.redoStack.push(redoActions);
    }
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  redo() {
    const actions = this.redoStack.pop();
    if (actions) {
      const undoActions: IAction[] = [];
      actions.forEach((action) => {
        action.redo();
        undoActions.push(action);
      });
      this.undoStack.push(undoActions);
    }
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }
}

export interface IAction {
  undo(): void;
  redo(): void;
}

type Cell = {
  sheet: number;
  row: number;
  column: number;
};

class ModelAction implements IAction {
  protected model: IModel;

  constructor(model: IModel) {
    this.model = model;
  }

  // eslint-disable-next-line class-methods-use-this
  protected redoAction() {
    throw new Error('Not implemented');
  }

  // eslint-disable-next-line class-methods-use-this
  protected undoAction() {
    throw new Error('Not implemented');
  }

  redo() {
    this.model.disableHistory();
    this.redoAction();
    this.model.enableHistory();
  }

  undo() {
    this.model.disableHistory();
    this.undoAction();
    this.model.enableHistory();
  }
}

export class SetCellValueAction extends ModelAction {
  private cell: Cell;

  private value: string | null;

  private oldValue: string | null;

  constructor(model: IModel, cell: Cell, value: string | null, oldValue: string | null) {
    super(model);

    this.cell = cell;
    this.value = value;
    this.oldValue = oldValue;
  }

  private setOrDeleteCellValue(value: string | null) {
    const { sheet, row, column } = this.cell;
    if (value) {
      this.model.setCellValue(sheet, row, column, value);
    } else {
      this.model.deleteCells(sheet, {
        rowStart: row,
        columnStart: column,
        rowEnd: row,
        columnEnd: column,
      });
    }
  }

  protected redoAction() {
    this.setOrDeleteCellValue(this.value);
  }

  protected undoAction() {
    this.setOrDeleteCellValue(this.oldValue);
  }
}

export class SetCellStyleAction extends ModelAction {
  private cell: Cell;

  private style: CellStyleSnapshot | null;

  private oldStyle: CellStyleSnapshot | null;

  constructor(
    model: IModel,
    cell: Cell,
    style: CellStyleSnapshot | null,
    oldStyle: CellStyleSnapshot | null,
  ) {
    super(model);

    this.cell = cell;
    this.style = style;
    this.oldStyle = oldStyle;
  }

  private setCellStyle(value: CellStyleSnapshot | null) {
    if (value === null) {
      return;
    }
    const { sheet, row, column } = this.cell;
    this.model.setCellsStyle(
      sheet,
      { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
      () => value,
    );
  }

  protected redoAction() {
    this.setCellStyle(this.style);
  }

  protected undoAction() {
    this.setCellStyle(this.oldStyle);
  }
}

export class InsertColumnAction extends ModelAction {
  private sheet: number;

  private column: number;

  constructor(model: IModel, sheet: number, column: number) {
    super(model);
    this.sheet = sheet;
    this.column = column;
  }

  protected redoAction() {
    this.model.insertColumn(this.sheet, this.column);
  }

  protected undoAction() {
    this.model.deleteColumn(this.sheet, this.column);
  }
}

export class SetColumnWidthAction extends ModelAction {
  private sheet: number;

  private column: number;

  private oldWidth: number;

  private width: number;

  constructor(model: IModel, sheet: number, column: number, oldWidth: number, width: number) {
    super(model);

    this.sheet = sheet;
    this.column = column;
    this.width = width;
    this.oldWidth = oldWidth;
  }

  private setColumnWidth(value: number) {
    this.model.setColumnWidth(this.sheet, this.column, value);
  }

  protected redoAction() {
    this.setColumnWidth(this.width);
  }

  protected undoAction() {
    this.setColumnWidth(this.oldWidth);
  }
}

export class DeleteColumnAction extends ModelAction {
  private sheet: number;

  private column: number;

  private width: number;

  private deletedCells: DeletedCell[];

  constructor(
    model: IModel,
    sheet: number,
    column: number,
    width: number,
    deletedCells: DeletedCell[],
  ) {
    super(model);
    this.sheet = sheet;
    this.column = column;
    this.width = width;
    this.deletedCells = deletedCells;
  }

  protected redoAction() {
    this.model.deleteColumn(this.sheet, this.column);
  }

  protected undoAction() {
    this.model.insertColumn(this.sheet, this.column);
    this.model.setColumnWidth(this.sheet, this.column, this.width);

    this.deletedCells.forEach((deletedCell) => {
      const { cell, value, style } = deletedCell;
      const { sheet, row, column } = cell;
      if (value) {
        this.model.setCellValue(sheet, row, column, value);
      }
      this.model.setCellsStyle(
        sheet,
        { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
        () => style,
      );
    });
  }
}

export class InsertRowAction extends ModelAction {
  private sheet: number;

  private row: number;

  constructor(model: IModel, sheet: number, row: number) {
    super(model);
    this.sheet = sheet;
    this.row = row;
  }

  protected redoAction() {
    this.model.insertRow(this.sheet, this.row);
  }

  protected undoAction() {
    this.model.deleteRow(this.sheet, this.row);
  }
}

export class SetRowHeightAction extends ModelAction {
  private sheet: number;

  private row: number;

  private oldHeight: number;

  private height: number;

  constructor(model: IModel, sheet: number, row: number, oldHeight: number, height: number) {
    super(model);

    this.sheet = sheet;
    this.row = row;
    this.height = height;
    this.oldHeight = oldHeight;
  }

  private setRowHeight(value: number) {
    this.model.setRowHeight(this.sheet, this.row, value);
  }

  protected redoAction() {
    this.setRowHeight(this.height);
  }

  protected undoAction() {
    this.setRowHeight(this.oldHeight);
  }
}

export class DeleteRowAction extends ModelAction {
  private sheet: number;

  private row: number;

  private height: number;

  private deletedCells: DeletedCell[];

  constructor(
    model: IModel,
    sheet: number,
    row: number,
    height: number,
    deletedCells: DeletedCell[],
  ) {
    super(model);
    this.sheet = sheet;
    this.row = row;
    this.height = height;
    this.deletedCells = deletedCells;
  }

  protected redoAction() {
    this.model.deleteRow(this.sheet, this.row);
  }

  protected undoAction() {
    this.model.insertRow(this.sheet, this.row);
    this.model.setRowHeight(this.sheet, this.row, this.height);

    this.deletedCells.forEach((deletedCell) => {
      const { cell, value, style } = deletedCell;
      const { sheet, row, column } = cell;
      if (value) {
        this.model.setCellValue(sheet, row, column, value);
      }
      this.model.setCellsStyle(
        sheet,
        { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
        () => style,
      );
    });
  }
}

export class AddBlankSheetAction extends ModelAction {
  private sheet: number;

  constructor(model: IModel, sheet: number) {
    super(model);

    this.sheet = sheet;
  }

  protected redoAction() {
    const sheet = this.model.addBlankSheet();
    this.sheet = sheet.index; // TODO: migrate redoStack to use new id
  }

  protected undoAction() {
    this.model.deleteSheet(this.sheet);
  }
}

export class RenameSheetAction extends ModelAction {
  private sheet: number;

  private name: string;

  private oldName: string;

  constructor(model: IModel, sheet: number, name: string, oldName: string) {
    super(model);

    this.sheet = sheet;
    this.name = name;
    this.oldName = oldName;
  }

  protected redoAction() {
    this.model.renameSheet(this.sheet, this.name);
  }

  protected undoAction() {
    this.model.renameSheet(this.sheet, this.oldName);
  }
}

export class DeleteSheetAction extends ModelAction {
  private sheet: number;

  private name: string;

  private deletedCells: DeletedCell[];

  constructor(model: IModel, sheet: number, name: string, deletedCells: DeletedCell[]) {
    super(model);

    this.sheet = sheet;
    this.name = name;
    this.deletedCells = deletedCells;
  }

  protected redoAction() {
    this.model.deleteSheet(this.sheet);
  }

  protected undoAction() {
    const sheet = this.model.addBlankSheet();
    if (sheet.name !== this.name) {
      sheet.name = this.name;
    }
    this.sheet = sheet.index; // TODO: Migrate undoStack to use the new id

    this.deletedCells.forEach((deletedCell) => {
      const { cell, value, style } = deletedCell;
      const { row, column } = cell;
      if (value) {
        this.model.setCellValue(this.sheet, row, column, value);
      }
      this.model.setCellsStyle(
        this.sheet,
        { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
        () => style,
      );
    });
  }
}
