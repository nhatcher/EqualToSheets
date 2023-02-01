import { useCallback, RefObject, ClipboardEvent } from 'react';
import { Action, WorkbookActionType } from './useWorkbookReducer';
import WorksheetCanvas from './canvas';
import { Area, Cell, NavigationKey, ScrollPosition } from './util';
import { Border } from './useKeyboardNavigation';
import Model from './model';
import { EditorSelection } from './editor/util';

export type WorkbookActions = {
  setAreaSelected: (area: Area, border: Border) => void;
  setScrollPosition: (position: ScrollPosition) => void;
  onSheetSelected: (sheet: number) => void;
  onAreaSelected: (area: Area, border: Border) => void;
  onCellsDeleted: () => void;
  onArrowUp: () => void;
  onArrowLeft: () => void;
  onArrowRight: () => void;
  onArrowDown: () => void;
  onKeyEnd: () => void;
  onKeyHome: () => void;
  onPageDown: () => void;
  onPageUp: () => void;
  onExtendToCell: (cell: Cell) => void;
  onExtendToEnd: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onToggleBold: () => void;
  onToggleItalic: () => void;
  onToggleUnderline: () => void;
  onToggleStrike: () => void;
  onToggleAlignLeft: () => void;
  onToggleAlignCenter: () => void;
  onToggleAlignRight: () => void;
  onTextColorPicked: (color: string) => void;
  onFillColorPicked: (color: string) => void;
  onAddBlankSheet: () => void;
  onSheetColorChanged: (color: string) => void;
  onSheetRenamed: (name: string) => void;
  onSheetDeleted: () => void;
  onNumberFormatPicked: (numberFmt: string) => void;
  onColumnWidthChanges: (sheet: number, column: number, width: number) => void;
  onRowHeightChanges: (sheet: number, row: number, height: number) => void;
  onDeleteRow: (sheet: number, row: number) => void;
  onInsertRow: (sheet: number, row: number) => void;
  onNavigationToEdge: (key: NavigationKey) => void;
  onEditKeyPressStart: (initText: string) => void;
  onExpandAreaSelectedKeyboard: (key: 'ArrowRight' | 'ArrowLeft' | 'ArrowDown' | 'ArrowUp') => void;
  onPointerMoveToCell: (cell: Cell) => void;
  onPointerDownAtCell: (cell: Cell, event: React.MouseEvent) => void;
  onEditChange: (text: string, cursorStart: number, cursorEnd: number) => void;
  onReferenceCycle: (text: string, cursorStart: number, cursorEnd: number) => void;
  onEditEnd: (options: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onCellEditStart: () => void;
  onBold: () => void;
  onItalic: () => void;
  onUnderline: () => void;
  onFormulaEditStart: (selection: EditorSelection) => void;
  onCopy: (event: ClipboardEvent) => void;
  onPaste: (event: ClipboardEvent) => void;
  onCut: (event: ClipboardEvent) => void;
  requestRender: () => void;
  resetModel: (model: Model) => void;
  focusWorkbook: () => void;
};

const useWorkbookActions = (
  dispatch: (value: Action) => void,
  {
    worksheetCanvas,
    worksheetElement,
    rootElement,
    onResize,
  }: {
    worksheetCanvas: RefObject<WorksheetCanvas>;
    worksheetElement: RefObject<HTMLDivElement>;
    rootElement: RefObject<HTMLDivElement>;
    onResize: (options: { deltaWidth: number; deltaHeight: number }) => void;
  },
): WorkbookActions => {
  const setAreaSelected = useCallback(
    (area: Area, border: Border): void => {
      dispatch({
        type: WorkbookActionType.SET_AREA_SELECTED,
        payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement, area, border },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const setScrollPosition = useCallback(
    (position: ScrollPosition): void => {
      dispatch({ type: WorkbookActionType.SET_SCROLL_POSITION, payload: { ...position } });
    },
    [dispatch],
  );

  const onSheetSelected = useCallback(
    (sheet: number): void => {
      dispatch({ type: WorkbookActionType.SELECT_SHEET, payload: { sheet } });
    },
    [dispatch],
  );

  const onAreaSelected = useCallback(
    (area: Area, border: Border): void => {
      dispatch({
        type: WorkbookActionType.AREA_SELECTED,
        payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement, area, border },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  // FIXME: Delete cells only works for areas [<top,left>, <bottom,right>]
  const onCellsDeleted = useCallback((): void => {
    dispatch({ type: WorkbookActionType.DELETE_CELLS });
  }, [dispatch]);

  const onArrowUp = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.ARROW_UP,
      payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onArrowLeft = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.ARROW_LEFT,
      payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onArrowRight = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.ARROW_RIGHT,
      payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onArrowDown = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.ARROW_DOWN,
      payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onKeyEnd = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.KEY_END,
      payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onKeyHome = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.KEY_HOME,
      payload: { canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onPageDown = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.PAGE_DOWN,
      payload: { worksheetRef: worksheetElement, canvasRef: worksheetCanvas },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onPageUp = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.PAGE_UP,
      payload: { worksheetRef: worksheetElement, canvasRef: worksheetCanvas },
    });
  }, [dispatch, worksheetCanvas, worksheetElement]);

  const onExtendToCell = useCallback(
    (cell: Cell): void => {
      dispatch({
        type: WorkbookActionType.EXTEND_TO_CELL,
        payload: { cell, worksheetRef: worksheetElement, canvasRef: worksheetCanvas },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const onExtendToEnd = useCallback((): void => {
    dispatch({ type: WorkbookActionType.EXTEND_TO_END });
  }, [dispatch]);

  const onUndo = useCallback((): void => {
    dispatch({ type: WorkbookActionType.UNDO });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onRedo = useCallback((): void => {
    dispatch({ type: WorkbookActionType.REDO });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleBold = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_BOLD });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleItalic = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_ITALIC });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleUnderline = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_UNDERLINE });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleStrike = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_STRIKE });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleAlignLeft = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_ALIGN_LEFT });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleAlignCenter = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_ALIGN_CENTER });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onToggleAlignRight = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_ALIGN_RIGHT });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onTextColorPicked = useCallback(
    (color: string): void => {
      dispatch({ type: WorkbookActionType.TEXT_COLOR_PICKED, payload: { color } });
      rootElement.current?.focus();
    },
    [dispatch, rootElement],
  );

  const onFillColorPicked = useCallback(
    (color: string): void => {
      dispatch({ type: WorkbookActionType.FILL_COLOR_PICKED, payload: { color } });
      rootElement.current?.focus();
    },
    [dispatch, rootElement],
  );

  const onSheetColorChanged = useCallback(
    (color: string): void => {
      dispatch({ type: WorkbookActionType.SHEET_COLOR_CHANGED, payload: { color } });
    },
    [dispatch],
  );

  const onSheetDeleted = useCallback((): void => {
    dispatch({ type: WorkbookActionType.SHEET_DELETED });
  }, [dispatch]);

  const onSheetRenamed = useCallback(
    (newName: string): void => {
      dispatch({ type: WorkbookActionType.SHEET_RENAMED, payload: { newName } });
    },
    [dispatch],
  );

  const onAddBlankSheet = useCallback((): void => {
    dispatch({ type: WorkbookActionType.ADD_BLANK_SHEET });
  }, [dispatch]);

  const onNumberFormatPicked = useCallback(
    (numberFmt: string): void => {
      dispatch({ type: WorkbookActionType.NUMBER_FORMAT_PICKED, payload: { numFmt: numberFmt } });
      rootElement.current?.focus();
    },
    [dispatch, rootElement],
  );

  const onColumnWidthChanges = useCallback(
    (sheet: number, column: number, width: number): void => {
      dispatch({
        type: WorkbookActionType.CHANGE_COLUMN_WIDTH,
        payload: { sheet, column, width, onResize },
      });
    },
    [dispatch, onResize],
  );

  const onRowHeightChanges = useCallback(
    (sheet: number, row: number, height: number): void => {
      dispatch({
        type: WorkbookActionType.CHANGE_ROW_HEIGHT,
        payload: { sheet, row, height, onResize },
      });
    },
    [dispatch, onResize],
  );

  const onDeleteRow = useCallback(
    (sheet: number, row: number): void => {
      dispatch({ type: WorkbookActionType.DELETE_ROW, payload: { sheet, row } });
    },
    [dispatch],
  );

  const onInsertRow = useCallback(
    (sheet: number, row: number): void => {
      dispatch({ type: WorkbookActionType.INSERT_ROW, payload: { sheet, row } });
    },
    [dispatch],
  );

  const onNavigationToEdge = useCallback(
    (key: NavigationKey): void => {
      dispatch({
        type: WorkbookActionType.NAVIGATION_TO_EDGE,
        payload: { key, worksheetRef: worksheetElement, canvasRef: worksheetCanvas },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const onEditKeyPressStart = useCallback(
    (initText: string): void => {
      dispatch({
        type: WorkbookActionType.EDIT_KEY_PRESS_START,
        payload: { initText },
      });
    },
    [dispatch],
  );

  const onExpandAreaSelectedKeyboard = useCallback(
    (key: 'ArrowRight' | 'ArrowLeft' | 'ArrowDown' | 'ArrowUp'): void => {
      dispatch({
        type: WorkbookActionType.EXPAND_AREA_SELECTED_KEYBOARD,
        payload: { key, canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const onPointerMoveToCell = useCallback(
    (cell: Cell): void => {
      dispatch({
        type: WorkbookActionType.POINTER_MOVE_TO_CELL,
        payload: { cell, canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const onPointerDownAtCell = useCallback(
    (cell: Cell, event: React.MouseEvent): void => {
      // See https://reactjs.org/docs/legacy-event-pooling.html
      event.persist();
      dispatch({
        type: WorkbookActionType.POINTER_DOWN_AT_CELL,
        payload: { cell, event, canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const onEditChange = useCallback(
    (text: string, cursorStart: number, cursorEnd: number): void => {
      dispatch({
        type: WorkbookActionType.EDIT_CHANGE,
        payload: { text, cursorStart, cursorEnd },
      });
    },
    [dispatch],
  );

  const onReferenceCycle = useCallback(
    (text: string, cursorStart: number, cursorEnd: number): void => {
      dispatch({
        type: WorkbookActionType.EDIT_REFERENCE_CYCLE,
        payload: { text, cursorStart, cursorEnd },
      });
    },
    [dispatch],
  );

  const onEditEnd = useCallback(
    (options: { deltaRow: number; deltaColumn: number }): void => {
      const { deltaRow, deltaColumn } = options;
      dispatch({
        type: WorkbookActionType.EDIT_END,
        payload: { deltaRow, deltaColumn, canvasRef: worksheetCanvas },
      });
      rootElement.current?.focus();
    },
    [dispatch, rootElement, worksheetCanvas],
  );

  const onEditEscape = useCallback((): void => {
    dispatch({ type: WorkbookActionType.EDIT_ESCAPE });
    rootElement.current?.focus();
  }, [dispatch, rootElement]);

  const onCellEditStart = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.EDIT_CELL_EDITOR_START,
      payload: {},
    });
  }, [dispatch]);

  const onBold = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_BOLD });
  }, [dispatch]);

  const onItalic = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_ITALIC });
  }, [dispatch]);

  const onUnderline = useCallback((): void => {
    dispatch({ type: WorkbookActionType.TOGGLE_UNDERLINE });
  }, [dispatch]);

  const onFormulaEditStart = useCallback(
    (selection: EditorSelection): void => {
      dispatch({
        type: WorkbookActionType.EDIT_FORMULA_BAR_EDITOR_START,
        payload: { selection },
      });
    },
    [dispatch],
  );

  // FIXME: Copy only works for areas [<top,left>, <bottom,right>]
  const onCopy = useCallback(
    (event: ClipboardEvent): void => {
      // See https://reactjs.org/docs/legacy-event-pooling.html
      event.persist();
      dispatch({ type: WorkbookActionType.COPY, payload: { event } });
      event.preventDefault();
      event.stopPropagation();
    },
    [dispatch],
  );

  const onPaste = useCallback(
    (event: ClipboardEvent): void => {
      // See https://reactjs.org/docs/legacy-event-pooling.html
      event.persist();
      const { items } = event.clipboardData;
      if (!items) {
        return;
      }
      const mimeTypes = ['application/json', 'text/plain', 'text/csv', 'text/html'];
      let mimeType;
      let value;
      const l = mimeTypes.length;
      for (let index = 0; index < l; index += 1) {
        mimeType = mimeTypes[index];
        value = event.clipboardData.getData(mimeType);
        if (value) {
          break;
        }
      }
      if (!mimeType || !value) {
        // No clipboard data to paste
        return;
      }
      dispatch({
        type: WorkbookActionType.PASTE,
        payload: { data: value, format: mimeType, event },
      });
      event.preventDefault();
      event.stopPropagation();
    },
    [dispatch],
  );

  const onCut = useCallback(
    (event: ClipboardEvent): void => {
      // See https://reactjs.org/docs/legacy-event-pooling.html
      event.persist();
      dispatch({ type: WorkbookActionType.CUT, payload: { event } });
      event.preventDefault();
      event.stopPropagation();
    },
    [dispatch],
  );

  const requestRender = useCallback(() => {
    dispatch({ type: WorkbookActionType.REQUEST_CANVAS_RENDER });
  }, [dispatch]);

  const resetModel = useCallback(
    (model: Model) => {
      dispatch({
        type: WorkbookActionType.RESET,
        payload: { model },
      });
    },
    [dispatch],
  );

  const focusWorkbook = useCallback(() => {
    rootElement.current?.focus();
  }, [rootElement]);

  return {
    setAreaSelected,
    setScrollPosition,
    onSheetSelected,
    onAreaSelected,
    onCellsDeleted,
    onArrowUp,
    onArrowLeft,
    onArrowRight,
    onArrowDown,
    onKeyEnd,
    onKeyHome,
    onPageDown,
    onPageUp,
    onExtendToCell,
    onExtendToEnd,
    onUndo,
    onRedo,
    onToggleBold,
    onToggleItalic,
    onToggleUnderline,
    onToggleStrike,
    onToggleAlignLeft,
    onToggleAlignCenter,
    onToggleAlignRight,
    onTextColorPicked,
    onFillColorPicked,
    onAddBlankSheet,
    onSheetColorChanged,
    onSheetRenamed,
    onSheetDeleted,
    onNumberFormatPicked,
    onColumnWidthChanges,
    onRowHeightChanges,
    onInsertRow,
    onDeleteRow,
    onNavigationToEdge,
    onEditKeyPressStart,
    onExpandAreaSelectedKeyboard,
    onPointerMoveToCell,
    onPointerDownAtCell,
    onEditChange,
    onReferenceCycle,
    onEditEnd,
    onEditEscape,
    onCellEditStart,
    onBold,
    onItalic,
    onUnderline,
    onFormulaEditStart,
    onCopy,
    onPaste,
    onCut,
    requestRender,
    resetModel,
    focusWorkbook,
  };
};

export default useWorkbookActions;
