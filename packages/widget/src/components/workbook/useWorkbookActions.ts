import { useCallback, RefObject } from 'react';
import { Action, WorkbookActionType } from './useWorkbookReducer/common';
import WorksheetCanvas from './canvas';
import { Area, Cell, NavigationKey, ScrollPosition, Border } from './util';

export type WorkbookActions = {
  setScrollPosition: (position: ScrollPosition) => void;
  onSheetSelected: (sheet: number) => void;
  onAreaSelected: (area: Area, border: Border) => void;
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
  onNavigationToEdge: (key: NavigationKey) => void;
  onEditKeyPressStart: (initText: string) => void;
  onSwitchCellEditorToEditMode: () => void;
  onExpandAreaSelectedKeyboard: (key: 'ArrowRight' | 'ArrowLeft' | 'ArrowDown' | 'ArrowUp') => void;
  onPointerMoveToCell: (cell: Cell) => void;
  onPointerDownAtCell: (cell: Cell) => void;
  onEditPointerDown: (cell: Cell, value: string) => void;
  onEditEnd: (text: string, delta: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onCellEditStart: () => void;
  onFormulaEditStart: () => void;
};

const useWorkbookActions = (
  dispatch: (value: Action) => void,
  {
    worksheetCanvas,
    worksheetElement,
    focusWorkbook,
  }: {
    worksheetCanvas: RefObject<WorksheetCanvas>;
    worksheetElement: RefObject<HTMLDivElement>;
    focusWorkbook: () => void;
  },
): WorkbookActions => {
  const setScrollPosition = useCallback(
    (position: ScrollPosition): void => {
      dispatch({ type: WorkbookActionType.SET_SCROLL_POSITION, payload: { ...position } });
    },
    [dispatch],
  );

  const onSheetSelected = useCallback(
    (sheet: number): void => {
      dispatch({ type: WorkbookActionType.SELECT_SHEET, payload: { sheet } });
      focusWorkbook();
    },
    [dispatch, focusWorkbook],
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

  const onSwitchCellEditorToEditMode = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.SWITCH_CELL_EDITOR_TO_EDIT_MODE,
      payload: {},
    });
  }, [dispatch]);

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

  const onEditPointerDown = useCallback(
    (cell: Cell, currentValue: string): void => {
      dispatch({
        type: WorkbookActionType.EDIT_POINTER_DOWN,
        payload: { cell, currentValue },
      });
    },
    [dispatch],
  );

  const onPointerDownAtCell = useCallback(
    (cell: Cell): void => {
      dispatch({
        type: WorkbookActionType.POINTER_DOWN_AT_CELL,
        payload: { cell, canvasRef: worksheetCanvas, worksheetRef: worksheetElement },
      });
    },
    [dispatch, worksheetCanvas, worksheetElement],
  );

  const onEditEnd = useCallback(
    (text: string, options: { deltaRow: number; deltaColumn: number }): void => {
      const { deltaRow, deltaColumn } = options;
      dispatch({
        type: WorkbookActionType.EDIT_END,
        payload: { deltaRow, deltaColumn, canvasRef: worksheetCanvas },
      });
      focusWorkbook();
    },
    [dispatch, focusWorkbook, worksheetCanvas],
  );

  const onEditEscape = useCallback((): void => {
    dispatch({ type: WorkbookActionType.EDIT_ESCAPE });
    focusWorkbook();
  }, [dispatch, focusWorkbook]);

  const onCellEditStart = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.EDIT_CELL_EDITOR_START,
      payload: {},
    });
  }, [dispatch]);

  const onFormulaEditStart = useCallback((): void => {
    dispatch({
      type: WorkbookActionType.EDIT_FORMULA_BAR_EDITOR_START,
      payload: {},
    });
  }, [dispatch]);

  return {
    setScrollPosition,
    onSheetSelected,
    onAreaSelected,
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
    onNavigationToEdge,
    onEditKeyPressStart,
    onSwitchCellEditorToEditMode,
    onExpandAreaSelectedKeyboard,
    onPointerMoveToCell,
    onPointerDownAtCell,
    onEditEnd,
    onEditEscape,
    onEditPointerDown,
    onCellEditStart,
    onFormulaEditStart,
  };
};

export default useWorkbookActions;
