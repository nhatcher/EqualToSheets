import Model from 'src/components/workbook/model';
import React, { RefObject, ClipboardEvent } from 'react';
import {
  StateSettings,
  ScrollPosition,
  Cell,
  Area,
  CellEditingType,
  AreaWithBorder,
  NavigationKey,
} from 'src/components/workbook/util';
import WorksheetCanvas from '../canvas';
import { Border } from '../useKeyboardNavigation';
import { EditorSelection } from '../editor/util';

export enum WorkbookLifecycle {
  Uninitialized,
  Initialized,
}

type WorkbookStateUninitialized = {
  lifecycle: WorkbookLifecycle.Uninitialized;
  // TODO: Maybe we could figure out a way to remove below props from the uninitialized state;
  model: null;
  selectedSheet: number;
  sheetStates: Record<string, unknown>;
  scrollPosition: ScrollPosition;
  selectedCell: Cell;
  selectedArea: Area;
  extendToArea: null;
  requestRenderId: number;
  formula: string;
  cellEditing: null;
  reducer: WorkbookReducer;
};

type WorkbookStateInitialized = {
  lifecycle: WorkbookLifecycle.Initialized;
  /** WARNING: This is kind of weird but model is mutable, only the reference is immutable.
   * Keeping the whole model immutable would slow performance a lot - we would need to pass a lot of data to Rust often.
   * We've decided to keep model inside the Reducer state even though it's mutable for the convenience.
   */
  model: Model;
  selectedSheet: number;
  sheetStates: {
    [sheetId: string]: StateSettings;
  };
  scrollPosition: ScrollPosition;
  selectedCell: Cell;
  /** NB the selected cell must be within the area selected */
  selectedArea: Area;
  extendToArea: AreaWithBorder | null;
  requestRenderId: number;
  formula: string;
  cellEditing: CellEditingType | null;
  reducer: WorkbookReducer;
};

export type WorkbookState = WorkbookStateInitialized | WorkbookStateUninitialized;

export type WorkbookReducer = (state: WorkbookState, action: Action) => WorkbookState;

export enum WorkbookActionType {
  /** Called when workbook json has changed, useful when you want to subscribe to the changes */
  WORKBOOK_HAS_CHANGED = 'WORKBOOK_HAS_CHANGED',
  /** Called when new model was created or workbookJson was replaced externally */
  RESET = 'RESET',
  SET_AREA_SELECTED = 'SET_AREA_SELECTED',
  SET_CELL_VALUE = 'SET_CELL_VALUE',
  EDIT_END = 'EDIT_END',
  EDIT_ESCAPE = 'EDIT_ESCAPE',
  EDIT_CHANGE = 'EDIT_CHANGE',
  EDIT_REFERENCE_CYCLE = 'EDIT_REFERENCE_CYCLE',
  EDIT_CELL_EDITOR_START = 'EDIT_CELL_EDITOR_START',
  EDIT_KEY_PRESS_START = 'EDIT_KEY_PRESS_START',
  EDIT_POINTER_MOVE = 'EDIT_POINTER_MOVE',
  EDIT_POINTER_DOWN = 'EDIT_POINTER_DOWN',
  EDIT_FORMULA_BAR_EDITOR_START = 'EDIT_FORMULA_BAR_EDITOR_START',
  SELECT_SHEET = 'SELECT_SHEET',
  SET_SCROLL_POSITION = 'SET_SCROLL_POSITION',
  ARROW_UP = 'ARROW_UP',
  ARROW_DOWN = 'ARROW_DOWN',
  ARROW_LEFT = 'ARROW_LEFT',
  ARROW_RIGHT = 'ARROW_RIGHT',
  KEY_END = 'KEY_END',
  KEY_HOME = 'KEY_HOME',
  CELL_SELECTED = 'CELL_SELECTED',
  PAGE_DOWN = 'PAGE_DOWN',
  PAGE_UP = 'PAGE_UP',
  AREA_SELECTED = 'AREA_SELECTED',
  POINTER_MOVE_TO_CELL = 'POINTER_MOVE_TO_CELL',
  POINTER_DOWN_AT_CELL = 'POINTER_DOWN_AT_CELL',
  EXPAND_AREA_SELECTED_POINTER = 'EXPAND_AREA_SELECTED_POINTER',
  EXPAND_AREA_SELECTED_KEYBOARD = 'EXPAND_AREA_SELECTED_KEYBOARD',
  EXTEND_TO_CELL = 'EXTEND_TO_CELL',
  EXTEND_TO_SELECTED = 'EXTEND_TO_SELECTED',
  EXTEND_TO_END = 'EXTEND_TO_END',
  UNDO = 'UNDO',
  REDO = 'REDO',
  DELETE_CELLS = 'DELETE_CELLS',
  TOGGLE_BOLD = 'TOGGLE_BOLD',
  TOGGLE_ITALIC = 'TOGGLE_ITALIC',
  TOGGLE_UNDERLINE = 'TOGGLE_UNDERLINE',
  TOGGLE_STRIKE = 'TOGGLE_STRIKE',
  TOGGLE_ALIGN_LEFT = 'TOGGLE_ALIGN_LEFT',
  TOGGLE_ALIGN_CENTER = 'TOGGLE_ALIGN_CENTER',
  TOGGLE_ALIGN_RIGHT = 'TOGGLE_ALIGN_RIGHT',
  TEXT_COLOR_PICKED = 'TEXT_COLOR_PICKED',
  FILL_COLOR_PICKED = 'FILL_COLOR_PICKED',
  SHEET_COLOR_CHANGED = 'SHEET_COLOR_CHANGED',
  SHEET_DELETED = 'SHEET_DELETED',
  SHEET_RENAMED = 'SHEET_RENAMED',
  ADD_BLANK_SHEET = 'ADD_BLANK_SHEET',
  NUMBER_FORMAT_PICKED = 'NUMBER_FORMAT_PICKED',
  CHANGE_COLUMN_WIDTH = 'CHANGE_COLUMN_WIDTH',
  CHANGE_ROW_HEIGHT = 'CHANGE_ROW_HEIGHT',
  DELETE_ROW = 'DELETE_ROW',
  INSERT_ROW = 'INSERT_ROW',
  NAVIGATION_TO_EDGE = 'NAVIGATION_TO_EDGE',
  COPY = 'COPY',
  PASTE = 'PASTE',
  CUT = 'CUT',
  SELECTED_CELL_HAS_CHANGED = 'SELECTED_CELL_HAS_CHANGED',
  REQUEST_CANVAS_RENDER = 'REQUEST_CANVAS_RENDER',
  CHANGE_REDUCER = 'CHANGE_REDUCER',
}

type PointerMoveToCellAction = {
  type: WorkbookActionType.POINTER_MOVE_TO_CELL;
  payload: {
    cell: Cell;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type SetCellValueAction = {
  type: WorkbookActionType.SET_CELL_VALUE;
  payload: {
    sheet: number;
    cell: Cell;
    text: string;
  };
};

type ExpandAreaSelectedPointerAction = {
  type: WorkbookActionType.EXPAND_AREA_SELECTED_POINTER;
  payload: {
    cell: Cell;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type EditPointerMoveAction = {
  type: WorkbookActionType.EDIT_POINTER_MOVE;
  payload: {
    cell: Cell;
  };
};

type PointerDownAtCellAction = {
  type: WorkbookActionType.POINTER_DOWN_AT_CELL;
  payload: {
    cell: Cell;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
    event: React.MouseEvent;
  };
};

type EditPointerDownAction = {
  type: WorkbookActionType.EDIT_POINTER_DOWN;
  payload: {
    cell: Cell;
  };
};

type ResetAction = {
  type: WorkbookActionType.RESET;
  payload: {
    model: Model;
  };
};

type EditEndAction = {
  type: WorkbookActionType.EDIT_END;
  payload: {
    deltaRow: number;
    deltaColumn: number;
    canvasRef: RefObject<WorksheetCanvas>;
  };
};

type EditEscapeAction = {
  type: WorkbookActionType.EDIT_ESCAPE;
};

type EditChangeAction = {
  type: WorkbookActionType.EDIT_CHANGE | WorkbookActionType.EDIT_REFERENCE_CYCLE;
  payload: {
    text: string;
    cursorStart: number;
    cursorEnd: number;
    /** Skips the readonly check */
    forceEdit?: boolean;
  };
};

type EditCellEditorStartAction = {
  type: WorkbookActionType.EDIT_CELL_EDITOR_START;
  payload: {
    /** Skips adding quote prefix to edited value */
    ignoreQuotePrefix?: boolean;
    /** Skips the readonly check */
    forceEdit?: boolean;
  };
};

type EditKeyPressStartAction = {
  type: WorkbookActionType.EDIT_KEY_PRESS_START;
  payload: {
    initText: string;
    /** Skips the readonly check */
    forceEdit?: boolean;
  };
};

type EditFormulaBarEditorStartAction = {
  type: WorkbookActionType.EDIT_FORMULA_BAR_EDITOR_START;
  payload: {
    selection: EditorSelection;
  };
};

type SelectSheetAction = {
  type: WorkbookActionType.SELECT_SHEET;
  payload: {
    sheet: number;
  };
};

type WorkbookHasChangedAction = {
  type: WorkbookActionType.WORKBOOK_HAS_CHANGED;
};

type UpdateModelAction = {
  type:
    | WorkbookActionType.EXTEND_TO_END
    | WorkbookActionType.UNDO
    | WorkbookActionType.REDO
    | WorkbookActionType.DELETE_CELLS
    | WorkbookActionType.TOGGLE_BOLD
    | WorkbookActionType.TOGGLE_ITALIC
    | WorkbookActionType.TOGGLE_UNDERLINE
    | WorkbookActionType.TOGGLE_STRIKE
    | WorkbookActionType.TOGGLE_ALIGN_LEFT
    | WorkbookActionType.TOGGLE_ALIGN_CENTER
    | WorkbookActionType.TOGGLE_ALIGN_RIGHT;
};

type ChangeColumnWidthAction = {
  type: WorkbookActionType.CHANGE_COLUMN_WIDTH;
  payload: {
    sheet: number;
    column: number;
    width: number;
    onResize: (options: { deltaWidth: number; deltaHeight: number }) => void;
  };
};

type ChangeRowHeightAction = {
  type: WorkbookActionType.CHANGE_ROW_HEIGHT;
  payload: {
    sheet: number;
    row: number;
    height: number;
    onResize: (options: { deltaWidth: number; deltaHeight: number }) => void;
  };
};

type DeleteRowAction = {
  type: WorkbookActionType.DELETE_ROW;
  payload: {
    sheet: number;
    row: number;
  };
};

type InsertRowAction = {
  type: WorkbookActionType.INSERT_ROW;
  payload: {
    sheet: number;
    row: number;
  };
};

type ColorUpdateAction = {
  type:
    | WorkbookActionType.TEXT_COLOR_PICKED
    | WorkbookActionType.FILL_COLOR_PICKED
    | WorkbookActionType.SHEET_COLOR_CHANGED;
  payload: {
    color: string;
  };
};

type DeleteSheetAction = {
  type: WorkbookActionType.SHEET_DELETED;
};
type RenameSheetAction = {
  type: WorkbookActionType.SHEET_RENAMED;
  payload: { newName: string };
};

type AddBlankSheetAction = {
  type: WorkbookActionType.ADD_BLANK_SHEET;
};

type NumberFmtUpdateAction = {
  type: WorkbookActionType.NUMBER_FORMAT_PICKED;
  payload: {
    numFmt: string;
  };
};

type SetScrollPositionAction = {
  type: WorkbookActionType.SET_SCROLL_POSITION;
  payload: {
    left: number;
    top: number;
  };
};

type SetAreaSelectedAction = {
  type: WorkbookActionType.SET_AREA_SELECTED;
  payload: {
    area: Area;
    border: Border;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type AreaSelectedAction = {
  type: WorkbookActionType.AREA_SELECTED;
  payload: {
    area: Area;
    border: Border;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type ExpandAreaSelectedKeyboardAction = {
  type: WorkbookActionType.EXPAND_AREA_SELECTED_KEYBOARD;
  payload: {
    key: 'ArrowRight' | 'ArrowLeft' | 'ArrowUp' | 'ArrowDown';
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type NavigationToEdgeAction = {
  type: WorkbookActionType.NAVIGATION_TO_EDGE;
  payload: {
    key: NavigationKey;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type CellSelectedAction = {
  type: WorkbookActionType.CELL_SELECTED;
  payload: {
    cell: Cell;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type ExtendToCellAction = {
  type: WorkbookActionType.EXTEND_TO_CELL;
  payload: {
    cell: Cell;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type ExtendToSelectedAction = {
  type: WorkbookActionType.EXTEND_TO_SELECTED;
  payload: {
    area: AreaWithBorder;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type CanvasAction = {
  type:
    | WorkbookActionType.ARROW_UP
    | WorkbookActionType.ARROW_DOWN
    | WorkbookActionType.ARROW_LEFT
    | WorkbookActionType.ARROW_RIGHT
    | WorkbookActionType.KEY_END
    | WorkbookActionType.KEY_HOME
    | WorkbookActionType.PAGE_DOWN
    | WorkbookActionType.PAGE_UP;
  payload: {
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
  };
};

type CopyAction = {
  type: WorkbookActionType.COPY;
  payload: {
    event: ClipboardEvent;
  };
};

type PasteAction = {
  type: WorkbookActionType.PASTE;
  payload: {
    format: string;
    data: string;
    event: ClipboardEvent;
  };
};

type CutAction = {
  type: WorkbookActionType.CUT;
  payload: {
    event: ClipboardEvent;
  };
};

type SelectedCellHasChangedAction = {
  type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED;
};

type RequestCanvasRenderAction = {
  type: WorkbookActionType.REQUEST_CANVAS_RENDER;
};

type ChangeReducerAction = {
  type: WorkbookActionType.CHANGE_REDUCER;
  payload: {
    reducer: WorkbookReducer;
  };
};

export type Action =
  | SetAreaSelectedAction
  | ChangeReducerAction
  | ResetAction
  | WorkbookHasChangedAction
  | SetCellValueAction
  | SelectedCellHasChangedAction
  | EditEndAction
  | SelectSheetAction
  | SetScrollPositionAction
  | CanvasAction
  | AreaSelectedAction
  | CellSelectedAction
  | ExtendToSelectedAction
  | ExtendToCellAction
  | UpdateModelAction
  | ChangeColumnWidthAction
  | ChangeRowHeightAction
  | DeleteRowAction
  | InsertRowAction
  | NavigationToEdgeAction
  | EditEscapeAction
  | EditChangeAction
  | EditCellEditorStartAction
  | EditFormulaBarEditorStartAction
  | CopyAction
  | PasteAction
  | CutAction
  | RequestCanvasRenderAction
  | EditKeyPressStartAction
  | ExpandAreaSelectedKeyboardAction
  | PointerMoveToCellAction
  | EditPointerMoveAction
  | ExpandAreaSelectedPointerAction
  | PointerDownAtCellAction
  | EditPointerDownAction
  | ColorUpdateAction
  | AddBlankSheetAction
  | DeleteSheetAction
  | RenameSheetAction
  | NumberFmtUpdateAction;
