import Model from 'src/components/workbook/model';
import { RefObject } from 'react';
import {
  Border,
  StateSettings,
  ScrollPosition,
  Cell,
  Area,
  CellEditingType,
  AreaWithBorder,
  NavigationKey,
} from 'src/components/workbook/util';
import WorksheetCanvas from '../canvas';

export type WorkbookState = {
  modelRef: React.RefObject<Model | null>;
  selectedSheet: number;
  sheetStates: {
    [sheetId: string]: StateSettings;
  };
  scrollPosition: ScrollPosition;
  selectedCell: Cell;
  /** NB the selected cell must be within the area selected */
  selectedArea: Area;
  extendToArea: AreaWithBorder | null;
  /** {@link CellEditingType} contains an ID. See comment there for more details. */
  cellEditingLastId: number;
  cellEditing: CellEditingType | null;
  reducer: WorkbookReducer;
};

export type WorkbookReducer = (state: WorkbookState, action: Action) => WorkbookState;

export enum WorkbookActionType {
  SET_CELL_VALUE = 'SET_CELL_VALUE',
  EDIT_END = 'EDIT_END',
  EDIT_ESCAPE = 'EDIT_ESCAPE',
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
  NAVIGATION_TO_EDGE = 'NAVIGATION_TO_EDGE',
}

type PointerMoveToCellAction = {
  type: WorkbookActionType.POINTER_MOVE_TO_CELL;
  payload: {
    cell: Cell;
    canvasRef: RefObject<WorksheetCanvas>;
    worksheetRef: RefObject<HTMLDivElement>;
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
  };
};

type EditPointerDownAction = {
  type: WorkbookActionType.EDIT_POINTER_DOWN;
  payload: {
    cell: Cell;
    currentValue: string;
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
  payload: {};
};

type SelectSheetAction = {
  type: WorkbookActionType.SELECT_SHEET;
  payload: {
    sheet: number;
  };
};

type UpdateModelAction = {
  type: WorkbookActionType.EXTEND_TO_END;
};

type SetScrollPositionAction = {
  type: WorkbookActionType.SET_SCROLL_POSITION;
  payload: {
    left: number;
    top: number;
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

export type Action =
  | EditEndAction
  | SelectSheetAction
  | SetScrollPositionAction
  | CanvasAction
  | AreaSelectedAction
  | CellSelectedAction
  | ExtendToSelectedAction
  | ExtendToCellAction
  | UpdateModelAction
  | NavigationToEdgeAction
  | EditEscapeAction
  | EditCellEditorStartAction
  | EditFormulaBarEditorStartAction
  | EditKeyPressStartAction
  | ExpandAreaSelectedKeyboardAction
  | PointerMoveToCellAction
  | EditPointerMoveAction
  | ExpandAreaSelectedPointerAction
  | PointerDownAtCellAction
  | EditPointerDownAction;
