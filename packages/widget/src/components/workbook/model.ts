import {
  CellStyleSnapshot,
  FormulaToken,
  ICell,
  ISheet,
  IWorkbook,
  NavigationDirection,
} from '@equalto-software/calc';
import Papa from 'papaparse';
import { TabsInput } from './components/navigation/common';
import { workbookLastColumn, workbookLastRow } from './constants';
import {
  ActionHistory,
  AddBlankSheetAction,
  DeleteColumnAction,
  DeleteRowAction,
  DeleteSheetAction,
  IAction,
  InsertColumnAction,
  InsertRowAction,
  RenameSheetAction,
  SetCellStyleAction,
  SetCellValueAction,
  SetColumnWidthAction,
  SetRowHeightAction,
} from './history';
import { IModel, StyleReducer } from './types';
import { Area, Cell, NavigationKey } from './util';

interface ModelSettings {
  workbook: IWorkbook;
  getTokens: (formula: string) => FormulaToken[];
  workbookFromJson: (workbookJson: string) => IWorkbook;
}

interface SheetArea extends Area {
  sheet: number;
}

export type CellStyle = ICell['style'];

interface CellData {
  value: string;
  style: CellStyleSnapshot;
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

type PasteType = 'copy' | 'cut';

// Replaces all tabs with spaces in a string
function escapeTabs(s: string): string {
  return s.replace(/\t/g, '  ');
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

// FIXME: We should use IDs instead of sheet indexes

export default class Model implements IModel {
  getTokens: (formula: string) => FormulaToken[];

  private workbookFromJson: (workbookJson: string) => IWorkbook;

  private history: ActionHistory;

  private workbook: IWorkbook;

  private nextSubscriberKey = 0;

  private subscribers: Record<number, Subscriber> = {};

  private subscriptionsPaused: boolean = false;

  constructor(options: ModelSettings) {
    this.workbook = options.workbook;
    this.history = new ActionHistory();
    this.getTokens = options.getTokens;
    this.workbookFromJson = options.workbookFromJson;
  }

  replaceWithJson(workbookJson: string) {
    this.workbook = this.workbookFromJson(workbookJson);
    this.notifySubscribers({ type: 'replaceWithJson' });
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
    if (this.subscriptionsPaused) {
      return;
    }
    for (const subscriber of Object.values(this.subscribers)) {
      subscriber(change);
    }
  }

  pauseSubscriptions() {
    this.subscriptionsPaused = true;
  }

  unpauseSubscriptions() {
    this.subscriptionsPaused = false;
  }

  setSheetColor(sheet: number, color: string): void {
    this.workbook.sheets.get(sheet).color = color;
    this.notifySubscribers({ type: 'setSheetColor' });
  }

  getColumnWidth(sheet: number, column: number): number {
    return this.workbook.sheets.get(sheet).getColumnWidth(column);
  }

  setColumnWidth(sheet: number, column: number, width: number): void {
    const oldWidth = this.getColumnWidth(sheet, column);

    this.workbook.sheets.get(sheet).setColumnWidth(column, width);

    const newWidth = this.getColumnWidth(sheet, column);
    this.history.push([new SetColumnWidthAction(this, sheet, column, oldWidth, newWidth)]);

    this.notifySubscribers({ type: 'setColumnWidth' });
  }

  getTabs(): TabsInput[] {
    return this.workbook.sheets.all().map((sheet) => {
      const { color } = sheet;
      return {
        index: sheet.index,
        sheet_id: sheet.id,
        name: sheet.name,
        color: color
          ? {
              RGB: color,
            }
          : undefined,
        state: 'visible',
      };
    });
  }

  getRowHeight(sheet: number, row: number): number {
    return this.workbook.sheets.get(sheet).getRowHeight(row);
  }

  setRowHeight(sheet: number, row: number, height: number): void {
    const oldHeight = this.getRowHeight(sheet, row);

    this.workbook.sheets.get(sheet).setRowHeight(row, height);

    const newHeight = this.getRowHeight(sheet, row);
    this.history.push([new SetRowHeightAction(this, sheet, row, oldHeight, newHeight)]);

    this.notifySubscribers({ type: 'setRowHeight' });
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

  addBlankSheet(): ISheet {
    const sheet = this.workbook.sheets.add();

    this.history.push([new AddBlankSheetAction(this, sheet.index)]);
    this.notifySubscribers({ type: 'addBlankSheet' });

    return sheet;
  }

  renameSheet(sheet: number, newName: string): void {
    const worksheet = this.workbook.sheets.get(sheet);
    const oldName = worksheet.name;

    worksheet.name = newName;

    this.history.push([new RenameSheetAction(this, sheet, newName, oldName)]);
    this.notifySubscribers({ type: 'renameSheet' });
  }

  deleteSheet(sheet: number): void {
    const worksheet = this.workbook.sheets.get(sheet);
    const { name } = worksheet;
    const deletedCells = worksheet.getCells().map((cell) => ({
      cell: {
        sheet,
        row: cell.row,
        column: cell.column,
      },
      value: Model.getInputValue(cell),
      style: cell.style.getSnapshot(),
    }));

    worksheet.delete();

    this.history.push([new DeleteSheetAction(this, sheet, name, deletedCells)]);
    this.notifySubscribers({ type: 'deleteSheet' });
  }

  setCellsStyle(sheet: number, area: Area, reducer: StyleReducer): void {
    const actions: IAction[] = [];
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
        const oldStyle = cell.style.getSnapshot();

        cell.style.bulkUpdate(reducer(cell.style));

        const newStyle = cell.style.getSnapshot();
        actions.push(new SetCellStyleAction(this, { sheet, row, column }, newStyle, oldStyle));
      }
    }
    this.history.push(actions);
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
    return this.getCellStyle(sheet, row, column).hasQuotePrefix;
  }

  getFormulaOrValue(sheet: number, row: number, column: number): string {
    const cell = this.workbook.cell(sheet, row, column);
    return cell.formula ?? `${cell.value ?? ''}`;
  }

  hasFormula(sheet: number, row: number, column: number): boolean {
    return this.workbook.cell(sheet, row, column).formula !== null;
  }

  enableHistory() {
    this.history.enable();
  }

  disableHistory() {
    this.history.disable();
  }

  canUndo(): boolean {
    return this.history.canUndo();
  }

  undo(): void {
    this.history.undo();
    this.notifySubscribers({ type: 'undo' });
  }

  canRedo(): boolean {
    return this.history.canRedo();
  }

  redo(): void {
    this.history.redo();
    this.notifySubscribers({ type: 'redo' });
  }

  private static getInputValue(cell: ICell): string | null {
    const oldValue = cell.formula ?? cell.value;
    if (oldValue === null) {
      return null;
    }
    return `${oldValue}`;
  }

  deleteCells(sheet: number, area: Area): void {
    const actions: IAction[] = [];

    for (let row = area.rowStart; row <= area.rowEnd; row += 1) {
      for (let column = area.columnStart; column <= area.columnEnd; column += 1) {
        const cell = this.workbook.cell(sheet, row, column);
        const oldValue = Model.getInputValue(cell);
        const oldStyle = cell.style.getSnapshot();

        // TODO: A lot of evaluations
        this.workbook.cell(sheet, row, column).value = null;
        actions.push(
          new SetCellValueAction(this, { sheet, row, column }, null, oldValue),
          new SetCellStyleAction(this, { sheet, row, column }, null, oldStyle),
        );
      }
    }

    this.history.push(actions);
    this.notifySubscribers({ type: 'deleteCells' });
  }

  setCellValue(sheet: number, row: number, column: number, value: string): void {
    const cell = this.workbook.cell(sheet, row, column);
    const oldValue = Model.getInputValue(cell);
    if (value === '') {
      // FIXME: This might be a bit of a HACK.
      // When deleting a cell we probably should call deleteCell and not setCellValue(null)
      // Although I don't think it is currently possible to set the value of a cell to be the empty string
      // To do that you would need to escape it with '
      cell.value = null;
      this.history.push([new SetCellValueAction(this, { sheet, row, column }, null, oldValue)]);
    } else {
      cell.input = value;
      this.history.push([new SetCellValueAction(this, { sheet, row, column }, value, oldValue)]);
    }
    this.notifySubscribers({ type: 'setCellValue' });
  }

  extendTo(sheet: number, sourceArea: Area, targetArea: Area): void {
    const actions: IAction[] = [];

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

        const cell = this.workbook.cell(sheet, targetRow, targetColumn);
        const oldValue = Model.getInputValue(cell);
        const oldStyle = cell.style.getSnapshot();

        this.extendToCell(sheet, sourceRow, sourceColumn, targetRow, targetColumn);

        const newValue = Model.getInputValue(cell);
        const newStyle = cell.style.getSnapshot();
        actions.push(
          new SetCellValueAction(
            this,
            { sheet, row: targetRow, column: targetColumn },
            newValue,
            oldValue,
          ),
          new SetCellStyleAction(
            this,
            { sheet, row: targetRow, column: targetColumn },
            newStyle,
            oldStyle,
          ),
        );
      }
    }
    this.history.push(actions);
    this.notifySubscribers({ type: 'extendTo' });
  }

  private extendToCell(
    sheet: number,
    sourceRow: number,
    sourceColumn: number,
    targetRow: number,
    targetColumn: number,
  ) {
    const sourceCell = this.workbook.cell(sheet, sourceRow, sourceColumn);
    const targetCell = this.workbook.cell(sheet, targetRow, targetColumn);

    const extendedValue = this.workbook.sheets
      .get(sheet)
      .userInterface.getExtendedValue(sourceRow, sourceColumn, targetRow, targetColumn);

    targetCell.input = extendedValue;
    targetCell.style = sourceCell.style;
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars, class-methods-use-this
  getFrozenRowsCount(sheet: number): number {
    return 0;
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars, class-methods-use-this
  getFrozenColumnsCount(sheet: number): number {
    return 0;
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
        const cell = this.workbook.cell(sheet, row, column);
        const value = cell.formula ?? `${cell.value ?? ''}`;
        const style = this.getCellStyle(sheet, row, column).getSnapshot();

        sheetData.data[row][column] = { value, style };
        tsvRow.push(escapeTabs(value));
      }

      tsv.push(tsvRow.join('\t'));
    }
    return { tsv: tsv.join('\n'), area, sheetData };
  }

  // This takes care of _internal_ paste, that is copy/paste within the application
  paste(source: SheetArea, target: SheetArea, sheetData: SheetData, type: PasteType): void {
    const actions: IAction[] = [];
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
        const cellData = sheetData.data[row][column];
        const targetRow = row + deltaRow;
        const targetColumn = column + deltaColumn;

        const cell = this.workbook.cell(target.sheet, targetRow, targetColumn);
        const oldValue = Model.getInputValue(cell);
        const oldStyle = cell.style.getSnapshot();
        const newStyle = cellData.style;
        let newValue;

        if (cellData.value === null) {
          newValue = '';
        } else {
          switch (type) {
            case 'copy': {
              newValue = this.workbook.getCopiedValueExtended(
                cellData.value,
                sheetData.sheetName,
                { sheet: source.sheet, row, column },
                { sheet: target.sheet, row: targetRow, column: targetColumn },
              );
              break;
            }
            case 'cut': {
              newValue = this.workbook.getCutValueMoved(
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

        cell.input = newValue;
        cell.style.bulkUpdate(newStyle);

        // TODO: Maybe way to stage actions?
        actions.push(
          new SetCellValueAction(
            this,
            { sheet: target.sheet, row: targetRow, column: targetColumn },
            newValue,
            oldValue,
          ),
          new SetCellStyleAction(
            this,
            { sheet: target.sheet, row: targetRow, column: targetColumn },
            cellData.style,
            oldStyle,
          ),
        );
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
      const { sheet } = source;
      for (let row = source.rowStart; row <= source.rowEnd; row += 1) {
        for (let column = source.columnStart; column <= source.columnEnd; column += 1) {
          if (!isCellInArea(source.sheet, row, column, targetArea)) {
            // FIXME: Should be able to reuse this.deleteCells
            const cell = this.workbook.cell(sheet, row, column);
            const oldValue = Model.getInputValue(cell);
            const oldStyle = cell.style.getSnapshot();

            // TODO: A lot of evaluations
            this.workbook.cell(sheet, row, column).value = null;
            actions.push(
              new SetCellValueAction(this, { sheet, row, column }, null, oldValue),
              new SetCellStyleAction(this, { sheet, row, column }, null, oldStyle),
            );
          }
        }
      }
      // Change all formulas that pointed to the old area to the new area
      // TODO: paste - we need forwardReferences diff getter
      const forwardReferencesActions = this.workbook.forwardReferences(sourceArea, {
        sheet: target.sheet,
        row: target.rowStart,
        column: target.columnStart,
      });
      actions.push(
        ...forwardReferencesActions.map((forwardReferenceAction) => {
          const { sheet: actionSheet, row, column } = forwardReferenceAction.cell;
          const { newValue, oldValue } = forwardReferenceAction;
          return new SetCellValueAction(
            this,
            { sheet: actionSheet, row, column },
            newValue,
            oldValue,
          );
        }),
      );
    }
    this.history.push(actions);
    this.notifySubscribers({ type: 'paste' });
  }

  pasteText(sheet: number, target: Cell, value: string): void {
    const actions: IAction[] = [];
    const parsedData = Papa.parse(value, { delimiter: '\t', header: false });
    const data = parsedData.data as string[][];

    for (let deltaRow = 0; deltaRow < data.length; deltaRow += 1) {
      const line = data[deltaRow];
      const targetRow = target.row + deltaRow;

      for (let deltaColumn = 0; deltaColumn < line.length; deltaColumn += 1) {
        const newValue = line[deltaColumn].trim();
        const targetColumn = target.column + deltaColumn;
        const cell = this.workbook.cell(sheet, targetRow, targetColumn);
        const oldValue = Model.getInputValue(cell);
        cell.input = newValue;
        actions.push(
          new SetCellValueAction(
            this,
            { sheet, row: targetRow, column: targetColumn },
            newValue,
            oldValue,
          ),
        );
      }
    }
    this.history.push(actions);
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

    this.history.push([new InsertRowAction(this, sheet, row)]);
    this.notifySubscribers({ type: 'insertRow' });
  }

  // We need to delete (and save) the row style
  // We need to delete (and save) the data
  deleteRow(sheet: number, row: number): void {
    const height = this.getRowHeight(sheet, row);
    const deletedCells = this.workbook.sheets
      .get(sheet)
      .getRowCells(row)
      .map((cell) => ({
        cell: {
          sheet,
          row: cell.row,
          column: cell.column,
        },
        value: Model.getInputValue(cell),
        style: cell.style.getSnapshot(),
      }));
    this.workbook.sheets.get(sheet).deleteRows(row, 1);

    this.history.push([new DeleteRowAction(this, sheet, row, height, deletedCells)]);
    this.notifySubscribers({ type: 'deleteRow' });
  }

  insertColumn(sheet: number, column: number): void {
    this.workbook.sheets.get(sheet).insertColumns(column, 1);

    this.history.push([new InsertColumnAction(this, sheet, column)]);
    this.notifySubscribers({ type: 'insertRow' });
  }

  // We need to delete (and save) the column style
  // We need to delete (and save) the data
  deleteColumn(sheet: number, column: number): void {
    const width = this.getColumnWidth(sheet, column);
    const deletedCells = this.workbook.sheets
      .get(sheet)
      .getColumnCells(column)
      .map((cell) => ({
        cell: {
          sheet,
          row: cell.row,
          column: cell.column,
        },
        value: Model.getInputValue(cell),
        style: cell.style.getSnapshot(),
      }));
    this.workbook.sheets.get(sheet).deleteColumns(column, 1);

    this.history.push([new DeleteColumnAction(this, sheet, column, width, deletedCells)]);
    this.notifySubscribers({ type: 'deleteColumn' });
  }

  saveToXlsx(): Uint8Array {
    return this.workbook.saveToXlsx();
  }

  toJson(): string {
    return this.workbook.toJson();
  }
}
