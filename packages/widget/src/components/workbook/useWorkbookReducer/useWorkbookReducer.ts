import { useReducer, Dispatch, useEffect } from 'react';
import {
  StateSettings,
  mergedAreas,
  FocusType,
  columnNameFromNumber,
  getExpandToArea,
  quoteSheetName,
} from 'src/components/workbook/util';
import {
  WorkbookActionType,
  Action,
  WorkbookLifecycle,
  WorkbookReducer,
  WorkbookState,
} from './common';
import { headerRowHeight, headerColumnWidth } from '../canvas';
import { Border } from '../useKeyboardNavigation';
import { getFormulaHTML, cycleReference, isInReferenceMode } from '../formulas';
import { CellStyle } from '../model';
import { getSelectedRangeInEditor } from '../editor/util';

export const CLIPBOARD_ID_SESSION_STORAGE_KEY = 'equalTo_clipboardId';

const getNewClipboardId = () => new Date().toISOString();

export const defaultSheetState: StateSettings = {
  scrollPosition: { left: 0, top: 0 },
  selectedCell: {
    column: 1,
    row: 1,
  },
  selectedArea: {
    rowStart: 1,
    columnStart: 1,
    rowEnd: 1,
    columnEnd: 1,
  },
  extendToArea: null,
};

export const getInitialWorkbookState = (reducer: WorkbookReducer): WorkbookState => ({
  lifecycle: WorkbookLifecycle.Uninitialized,
  model: null,
  tabs: [],
  selectedSheet: 0,
  sheetStates: {},
  requestRenderId: 0,
  formula: '',
  cellEditing: null,
  scrollPosition: { ...defaultSheetState.scrollPosition },
  selectedCell: { ...defaultSheetState.selectedCell },
  selectedArea: { ...defaultSheetState.selectedArea },
  extendToArea: null,
  reducer,
});

export const defaultWorkbookReducer: WorkbookReducer = (state, action): WorkbookState => {
  if (state.lifecycle === WorkbookLifecycle.Uninitialized) {
    switch (action.type) {
      case WorkbookActionType.RESET: {
        const { model: newModel, assignmentsJson } = action.payload;
        if (assignmentsJson) {
          newModel.setCellsWithValuesJsonIgnoreReadonly(assignmentsJson);
        }
        return state.reducer(
          {
            ...state,
            lifecycle: WorkbookLifecycle.Initialized,
            selectedSheet: 0,
            model: newModel,
            tabs: newModel.getTabs(),
            sheetStates: {},
            requestRenderId: 0,
            formula: '',
            cellEditing: null,
            // TODO: This is a bit hacky in that there are two states for the sheet flying around
            // this one and the one given by {scrollPosition, selectedCell, selectedArea} --Nico
            ...defaultSheetState,
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      default:
        // TODO: We should log these actions
        // Before workbook becomes initialized ignore all other actions
        return state;
    }
  } else {
    switch (action.type) {
      case WorkbookActionType.WORKBOOK_HAS_CHANGED: {
        const { model, selectedSheet, selectedCell, requestRenderId } = state;
        const newFormula = model.getFormulaOrValue(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        return {
          ...state,
          tabs: model.getTabs(),
          requestRenderId: requestRenderId + 1,
          formula: newFormula,
        };
      }

      case WorkbookActionType.CHANGE_REDUCER: {
        return {
          ...state,
          reducer: action.payload.reducer,
        };
      }

      case WorkbookActionType.REQUEST_CANVAS_RENDER: {
        return {
          ...state,
          requestRenderId: state.requestRenderId + 1,
        };
      }

      case WorkbookActionType.SELECTED_CELL_HAS_CHANGED: {
        const { model, selectedSheet, selectedCell } = state;
        const newFormula = model.getFormulaOrValue(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        return {
          ...state,
          formula: newFormula,
        };
      }

      case WorkbookActionType.RESET: {
        const { model: newModel, assignmentsJson } = action.payload;
        if (assignmentsJson) {
          newModel.setCellsWithValuesJsonIgnoreReadonly(assignmentsJson);
        }
        const { selectedSheet, requestRenderId, sheetStates } = state;

        // Save sheet state
        sheetStates[state.tabs[selectedSheet].sheet_id] = {
          scrollPosition: { ...state.scrollPosition },
          selectedCell: { ...state.selectedCell },
          selectedArea: { ...state.selectedArea },
          extendToArea: state.extendToArea && { ...state.extendToArea },
        };

        const newTabs = newModel.getTabs();
        const newSelectedSheet =
          selectedSheet < newTabs.length ? selectedSheet : newTabs.length - 1;

        const newSheetId = newTabs[newSelectedSheet].sheet_id;
        const sheetState = newSheetId in sheetStates ? sheetStates[newSheetId] : defaultSheetState;

        return state.reducer(
          {
            ...state,
            model: newModel,
            tabs: newTabs,
            selectedSheet: newSelectedSheet,
            requestRenderId: requestRenderId + 1,
            sheetStates: { ...sheetStates },
            selectedCell: { ...sheetState.selectedCell },
            selectedArea: { ...sheetState.selectedArea },
            scrollPosition: { ...sheetState.scrollPosition },
            extendToArea: sheetState.extendToArea && { ...sheetState.extendToArea },
            cellEditing: state.cellEditing && { ...state.cellEditing, focus: FocusType.FormulaBar },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.SET_ASSIGNMENTS: {
        const { model, selectedSheet, selectedCell, requestRenderId } = state;
        model.setCellsWithValuesJsonIgnoreReadonly(action.payload.assignmentsJson);
        const newFormula = model.getFormulaOrValue(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        return {
          ...state,
          requestRenderId: requestRenderId + 1,
          formula: newFormula,
        };
      }

      case WorkbookActionType.EDIT_END: {
        const { cellEditing } = state;
        if (!cellEditing) {
          return state;
        }
        const { deltaRow, deltaColumn, canvasRef } = action.payload;
        const canvas = canvasRef.current;
        if (!canvas) {
          return state;
        }
        const { sheet, row, column, text } = cellEditing;

        const newRow = row + deltaRow;
        const newColumn = column + deltaColumn;
        let newState = state;
        if (row >= 1 && row <= canvas.lastRow && column >= 1 && column <= canvas.lastColumn) {
          newState = {
            ...state,
            selectedSheet: sheet,
            selectedCell: { row: newRow, column: newColumn },
            selectedArea: {
              rowStart: newRow,
              rowEnd: newRow,
              columnStart: newColumn,
              columnEnd: newColumn,
            },
          };
        }

        return state.reducer(
          { ...newState, cellEditing: null },
          {
            type: WorkbookActionType.SET_CELL_VALUE,
            payload: {
              sheet,
              cell: { row, column },
              text,
            },
          },
        );
      }

      case WorkbookActionType.EDIT_ESCAPE: {
        return { ...state, cellEditing: null };
      }

      case WorkbookActionType.EDIT_CHANGE: {
        const { model, cellEditing } = state;
        if (!cellEditing) {
          return state;
        }
        const { sheet, row, column } = cellEditing;
        if (!action.payload.forceEdit && model.isCellReadOnly(sheet, row, column)) {
          return state;
        }
        const { text, cursorStart, cursorEnd } = action.payload;
        const { html, activeRanges } = getFormulaHTML(
          text,
          sheet,
          model.getTabs().map((tab) => tab.name),
          model.getTokens,
        );
        return {
          ...state,
          cellEditing: {
            ...cellEditing,
            text,
            html,
            activeRanges,
            cursorStart,
            cursorEnd,
          },
        };
      }

      case WorkbookActionType.EDIT_REFERENCE_CYCLE: {
        const { model, cellEditing } = state;
        if (!cellEditing) {
          return state;
        }
        const { sheet, row, column } = cellEditing;
        if (model.isCellReadOnly(sheet, row, column)) {
          return state;
        }
        const { text, cursorStart, cursorEnd } = action.payload;
        const newFormula = cycleReference({
          text,
          context: { sheet, row, column },
          getTokens: model.getTokens,
          cursor: {
            start: cursorStart,
            end: cursorEnd,
          },
        });
        const { html, activeRanges } = getFormulaHTML(
          newFormula.text,
          sheet,
          model.getTabs().map((tab) => tab.name),
          model.getTokens,
        );
        return {
          ...state,
          cellEditing: {
            ...cellEditing,
            text: newFormula.text,
            html,
            activeRanges,
            cursorStart: newFormula.cursorStart,
            cursorEnd: newFormula.cursorEnd,
          },
        };
      }

      case WorkbookActionType.EDIT_KEY_PRESS_START: {
        const { model, selectedSheet, selectedCell } = state;
        const { row, column } = selectedCell;
        if (!action.payload.forceEdit && model.isCellReadOnly(selectedSheet, row, column)) {
          return state;
        }
        const { initText } = action.payload;
        const html = `<span>${initText}</span>`;

        return {
          ...state,
          cellEditing: {
            sheet: selectedSheet,
            row: selectedCell.row,
            column: selectedCell.column,
            text: initText,
            base: initText,
            html,
            cursorStart: initText.length,
            cursorEnd: initText.length,
            mode: 'init',
            focus: FocusType.Cell,
            activeRanges: [],
          },
        };
      }

      case WorkbookActionType.EDIT_CELL_EDITOR_START: {
        const { model, selectedSheet, selectedCell, cellEditing } = state;
        if (cellEditing) {
          // Ignore if already editing
          return state;
        }
        const { row, column } = selectedCell;
        if (model.isCellReadOnly(selectedSheet, row, column)) {
          return state;
        }
        let text = model.getFormulaOrValue(selectedSheet, row, column);
        if (!action.payload.ignoreQuotePrefix && model.isQuotePrefix(selectedSheet, row, column)) {
          text = `'${text}`;
        }
        const { html, activeRanges } = getFormulaHTML(
          text,
          selectedSheet,
          model.getTabs().map((tab) => tab.name),
          model.getTokens,
        );
        const cursor = text.length;
        return {
          ...state,
          cellEditing: {
            text,
            base: text,
            focus: FocusType.Cell,
            sheet: selectedSheet,
            row,
            column,
            html,
            cursorStart: cursor,
            cursorEnd: cursor,
            mode: 'edit',
            activeRanges,
          },
        };
      }

      case WorkbookActionType.EDIT_FORMULA_BAR_EDITOR_START: {
        const { cellEditing } = state;
        if (cellEditing) {
          return state;
        }

        // Start completely new edit
        const { model, selectedCell, selectedSheet } = state;
        const { row, column } = selectedCell;
        if (model.isCellReadOnly(selectedSheet, row, column)) {
          return state;
        }
        let text = model.getFormulaOrValue(selectedSheet, row, column);
        if (model.isQuotePrefix(selectedSheet, row, column)) {
          text = `'${text}`;
        }
        const { html, activeRanges } = getFormulaHTML(
          text,
          selectedSheet,
          model.getTabs().map((tab) => tab.name),
          model.getTokens,
        );

        const { selection } = action.payload;

        return {
          ...state,
          cellEditing: {
            text,
            base: text,
            focus: FocusType.FormulaBar,
            sheet: selectedSheet,
            row,
            column,
            html,
            cursorStart: selection.start,
            cursorEnd: selection.end,
            mode: 'edit',
            activeRanges,
          },
        };
      }

      case WorkbookActionType.SELECT_SHEET: {
        const { sheetStates, tabs } = state;
        // Save sheet state

        sheetStates[state.tabs[state.selectedSheet].sheet_id] = {
          scrollPosition: { ...state.scrollPosition },
          selectedCell: { ...state.selectedCell },
          selectedArea: { ...state.selectedArea },
          extendToArea: state.extendToArea && { ...state.extendToArea },
        };
        let { sheet } = action.payload;

        // In case the sheet does not exist anymore (i.e. has been deleted)
        if (tabs.length <= sheet) {
          sheet = 0;
        }

        const sheetId = tabs[sheet].sheet_id;
        const sheetState = sheetId in sheetStates ? sheetStates[sheetId] : defaultSheetState;

        return state.reducer(
          {
            ...state,
            selectedSheet: sheet,
            sheetStates: { ...sheetStates },
            selectedCell: { ...sheetState.selectedCell },
            selectedArea: { ...sheetState.selectedArea },
            scrollPosition: { ...sheetState.scrollPosition },
            extendToArea: sheetState.extendToArea && { ...sheetState.extendToArea },
            cellEditing: state.cellEditing && { ...state.cellEditing, focus: FocusType.FormulaBar },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.SET_SCROLL_POSITION: {
        const { left, top } = action.payload;
        if (left === state.scrollPosition.left && top === state.scrollPosition.top) {
          return state;
        }
        return { ...state, scrollPosition: { left, top } };
      }

      case WorkbookActionType.ARROW_UP: {
        const { canvasRef } = action.payload;
        const canvas = canvasRef.current;

        const newRow = state.selectedCell.row - 1;
        if (!canvas || newRow < 1) {
          return state;
        }

        const newRowY = canvas.getCoordinatesByCell(newRow, state.selectedCell.column)[1];
        let scrollTop = state.scrollPosition.top;
        const frozenHeight = canvas.getFrozenRowsHeight();
        if (newRowY < headerRowHeight + frozenHeight) {
          const { top: canvasScrollTop } = canvas.getScrollPosition();
          scrollTop = canvasScrollTop - (headerRowHeight + frozenHeight - newRowY) + 1;
        }
        const { column } = state.selectedCell;

        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, row: newRow },
            selectedArea: {
              columnStart: column,
              columnEnd: column,
              rowStart: newRow,
              rowEnd: newRow,
            },
            scrollPosition: { ...state.scrollPosition, top: scrollTop },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.ARROW_LEFT: {
        const { canvasRef } = action.payload;
        const canvas = canvasRef.current;

        const newColumn = state.selectedCell.column - 1;
        const { row } = state.selectedCell;
        if (!canvas || newColumn < 1) {
          return state;
        }

        let scrollLeft = state.scrollPosition.left;
        const newColumnX = canvas.getCoordinatesByCell(state.selectedCell.row, newColumn)[0];
        const frozenWidth = canvas.getFrozenColumnsWidth();
        if (newColumnX < headerColumnWidth + frozenWidth) {
          const { left: canvasScrollLeft } = canvas.getScrollPosition();
          scrollLeft = canvasScrollLeft - (headerColumnWidth + frozenWidth - newColumnX) + 3;
        }

        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, column: newColumn },
            selectedArea: {
              rowStart: row,
              rowEnd: row,
              columnStart: newColumn,
              columnEnd: newColumn,
            },
            scrollPosition: { ...state.scrollPosition, left: scrollLeft },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.ARROW_RIGHT: {
        const { canvasRef } = action.payload;
        const canvas = canvasRef.current;

        const newColumn = state.selectedCell.column + 1;
        const { row } = state.selectedCell;
        if (!canvas || newColumn > canvas.lastColumn) {
          return state;
        }

        let scrollLeft = state.scrollPosition.left;
        const frozenColumns = canvas.workbook.getFrozenColumnsCount();
        // If we press right arrow in the last of the frozen columns,
        // we scroll back to the origin and go to the next column
        if (state.selectedCell.column === frozenColumns) {
          scrollLeft = 0;
        } else if (newColumn > frozenColumns) {
          const { width: canvasWidth } = canvas;
          const newColumnRightX = canvas.getCoordinatesByCell(
            state.selectedCell.row,
            newColumn + 1,
          )[0];
          if (newColumnRightX > canvasWidth - 10) {
            scrollLeft = canvas.getMinScrollLeft(newColumn) + 10;
          }
        }

        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, column: newColumn },
            selectedArea: {
              rowStart: row,
              rowEnd: row,
              columnStart: newColumn,
              columnEnd: newColumn,
            },
            scrollPosition: { ...state.scrollPosition, left: scrollLeft },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.ARROW_DOWN: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;

        const newRow = state.selectedCell.row + 1;
        const { column } = state.selectedCell;
        if (!canvas || !worksheet || newRow > canvas.lastRow) {
          return state;
        }

        let scrollTop = state.scrollPosition.top;

        // If we are the last of the frozen rows span to the origin
        if (state.selectedCell.row === canvas.workbook.getFrozenRowsCount()) {
          scrollTop = 0;
        } else {
          const { height } = worksheet.getBoundingClientRect();
          const newCellBottom = canvas.getCoordinatesByCell(
            newRow + 1,
            state.selectedCell.column,
          )[1];
          const approximateScrollWidth = 10;
          if (newCellBottom > height - approximateScrollWidth) {
            // If the cellBottom is closer to height than scrollbarWidth units we scroll
            const newCellTop = canvas.getCoordinatesByCell(newRow, state.selectedCell.column)[1];
            const { top: canvasTop } = canvas.getScrollPosition();
            const cellHeight = newCellBottom - newCellTop;
            // We need to scroll at least cellHeight. Sine we can only scroll in integer cells
            // we find the next cell whose top is larger than cellHeight
            const scrollToCell = canvas.getCellByCoordinates(100, headerRowHeight + cellHeight - 1);
            if (!scrollToCell) {
              return state;
            }
            // The bottom of 'scrollToCell' is the coordinates of the next row.
            const minY = canvas.getCoordinatesByCell(
              scrollToCell.row + 1,
              state.selectedCell.column,
            )[1];
            scrollTop = canvasTop + (minY - headerRowHeight) + 1;
          }
        }

        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, row: newRow },
            selectedArea: {
              rowStart: newRow,
              rowEnd: newRow,
              columnStart: column,
              columnEnd: column,
            },
            scrollPosition: { ...state.scrollPosition, top: scrollTop },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.KEY_END: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }
        const { row } = state.selectedCell;
        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, column: canvas.lastColumn },
            selectedArea: {
              ...state.selectedArea,
              columnStart: canvas.lastColumn,
              columnEnd: canvas.lastColumn,
              rowStart: row,
              rowEnd: row,
            },
            scrollPosition: { ...state.scrollPosition, left: canvas.sheetWidth },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.KEY_HOME: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }
        const { row } = state.selectedCell;
        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, column: 1 },
            selectedArea: { columnStart: 1, columnEnd: 1, rowStart: row, rowEnd: row },
            scrollPosition: { ...state.scrollPosition, left: 0 },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.PAGE_DOWN: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }

        const frozenRows = canvas.workbook.getFrozenRowsCount();
        if (state.selectedCell.row < frozenRows) {
          return state;
        }

        const { topLeftCell, bottomRightCell } = canvas.getVisibleCells();
        const delta = state.selectedCell.row - topLeftCell.row;
        const newRow = bottomRightCell.row + delta;
        const { column } = state.selectedCell;
        if (newRow > canvas.lastRow) {
          return state;
        }
        const bottomRightCellY = canvas.getCoordinatesByCell(
          bottomRightCell.row,
          bottomRightCell.column,
        )[1];
        const scrollTop =
          bottomRightCellY + canvas.getScrollPosition().top - canvas.getFrozenRowsHeight() - 4;

        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, row: newRow },
            selectedArea: {
              rowStart: newRow,
              rowEnd: newRow,
              columnStart: column,
              columnEnd: column,
            },
            scrollPosition: { ...state.scrollPosition, top: scrollTop },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.PAGE_UP: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }

        // Compute scroll - magic by Nico
        const { topLeftCell } = canvas.getVisibleCells();
        const delta = state.selectedCell.row - topLeftCell.row;
        const topCellRowHeight = state.model.getRowHeight(state.selectedSheet, topLeftCell.row);
        const deltaHeight = canvas.height - headerRowHeight - topCellRowHeight;
        const scrollTop = Math.max(0, canvas.getScrollPosition().top - deltaHeight);

        const newRow = canvas.getBoundedRow(scrollTop).row + delta;
        const { column } = state.selectedCell;

        return state.reducer(
          {
            ...state,
            selectedCell: { ...state.selectedCell, row: newRow },
            selectedArea: {
              ...state.selectedArea,
              rowStart: newRow,
              rowEnd: newRow,
              columnStart: column,
              columnEnd: column,
            },
            scrollPosition: { ...state.scrollPosition, top: scrollTop },
          },
          { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
        );
      }

      case WorkbookActionType.AREA_SELECTED: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }

        const { area, border } = action.payload;
        const { rowStart, rowEnd, columnStart, columnEnd } = area;

        if (
          rowStart > 0 &&
          columnStart > 0 &&
          rowEnd <= canvas.lastRow &&
          columnEnd <= canvas.lastColumn
        ) {
          // FIXME: Here and onCellSelected we scroll to the visible Cell
          // Scrolling should happen in multiples of column widths and row heights
          // We use 20 as a harmless guestimate of the scrollbar width.
          // Magic scrolling by Nico
          const { width, height } = worksheet.getBoundingClientRect();
          let { left, top } = state.scrollPosition;
          switch (border) {
            case Border.Right: {
              let x = canvas.getCoordinatesByCell(rowStart, columnEnd)[0];
              x += canvas.workbook.getColumnWidth(columnEnd);
              if (x > width - 20) {
                left += x - width + 10;
              }

              break;
            }
            case Border.Left: {
              const x = canvas.getCoordinatesByCell(rowStart, columnStart)[0];
              if (x < headerColumnWidth) {
                left -= headerColumnWidth - x;
              }

              break;
            }
            case Border.Top: {
              const y = canvas.getCoordinatesByCell(rowStart, columnStart)[1];
              if (y < headerRowHeight) {
                top -= headerRowHeight - y;
              }

              break;
            }
            case Border.Bottom: {
              const y = canvas.getCoordinatesByCell(rowEnd + 1, columnStart)[1];
              if (y > height - 20) {
                top += y - height + 10;
              }

              break;
            }
            default:
              break;
          }
          // If there are frozen rows or columns snap to origin if we cross boundaries
          const frozenRows = canvas.workbook.getFrozenRowsCount();
          const frozenColumns = canvas.workbook.getFrozenColumnsCount();
          if (area.rowStart <= frozenRows && area.rowEnd > frozenRows) {
            top = 0;
          }
          if (area.columnStart <= frozenColumns && area.columnEnd > frozenColumns) {
            left = 0;
          }
          return {
            ...state,
            selectedArea: { ...area },
            scrollPosition: { left, top },
          };
        }
        return state;
      }

      case WorkbookActionType.EXPAND_AREA_SELECTED_KEYBOARD: {
        const { row, column } = state.selectedCell;
        let { rowStart, rowEnd, columnStart, columnEnd } = state.selectedArea;
        const { key, canvasRef, worksheetRef } = action.payload;
        let border = Border.Left;

        switch (key) {
          case 'ArrowRight': {
            if (columnStart < column) {
              border = Border.Left;
              columnStart += 1;
            } else {
              border = Border.Right;
              columnEnd += 1;
            }

            break;
          }
          case 'ArrowLeft': {
            if (column < columnEnd) {
              border = Border.Right;
              columnEnd -= 1;
            } else {
              border = Border.Left;
              columnStart -= 1;
            }

            break;
          }
          case 'ArrowUp': {
            if (row < rowEnd) {
              border = Border.Bottom;
              rowEnd -= 1;
            } else {
              border = Border.Top;
              rowStart -= 1;
            }

            break;
          }
          case 'ArrowDown': {
            if (row > rowStart) {
              border = Border.Top;
              rowStart += 1;
            } else {
              border = Border.Bottom;
              rowEnd += 1;
            }

            break;
          }
          default:
            break;
        }

        return state.reducer(state, {
          type: WorkbookActionType.AREA_SELECTED,
          payload: {
            border,
            area: { rowStart, rowEnd, columnStart, columnEnd },
            canvasRef,
            worksheetRef,
          },
        });
      }

      case WorkbookActionType.POINTER_DOWN_AT_CELL: {
        const { cellEditing } = state;
        const { cell, canvasRef, worksheetRef, event } = action.payload;
        if (cellEditing && isInReferenceMode(cellEditing.text, cellEditing.cursorEnd)) {
          event.preventDefault();
          event.stopPropagation();
          return state.reducer(state, {
            type: WorkbookActionType.EDIT_POINTER_DOWN,
            payload: { cell },
          });
        }
        if (cellEditing) {
          // FIXME: This is out of context. This happens while you are editing a cell and finish the editing by
          // clicking somewhere else.
          // This probably should be done in the onBlur event in the editor.
          const { sheet, row, column } = cellEditing;
          // The cellEditing object might have not been updated because we debounce key strokes,
          // So we use the text in the editor (note that we are editing the cell)
          const sel = getSelectedRangeInEditor();
          const text = sel ? sel.text : '';

          const newRow = cell.row;
          const newColumn = cell.column;
          const newState = {
            ...state,
            selectedSheet: sheet,
            selectedCell: { row: newRow, column: newColumn },
            selectedArea: {
              rowStart: newRow,
              rowEnd: newRow,
              columnStart: newColumn,
              columnEnd: newColumn,
            },
          };

          return state.reducer(
            { ...newState, cellEditing: null },
            {
              type: WorkbookActionType.SET_CELL_VALUE,
              payload: {
                sheet,
                cell: { row, column },
                text,
              },
            },
          );
        }
        return state.reducer(state, {
          type: WorkbookActionType.CELL_SELECTED,
          payload: { cell, canvasRef, worksheetRef },
        });
      }

      case WorkbookActionType.SET_CELL_VALUE: {
        const { sheet, cell, text } = action.payload;
        state.model.setCellValue(sheet, cell.row, cell.column, text);
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.EDIT_POINTER_DOWN: {
        const { model, cellEditing, selectedSheet, tabs } = state;
        if (!cellEditing) {
          return state;
        }
        const { cell } = action.payload;
        const { row, column } = cell;
        const prefix =
          cellEditing.sheet === selectedSheet ? '' : `${quoteSheetName(tabs[selectedSheet].name)}!`;
        const text = `${cellEditing.text}${prefix}${columnNameFromNumber(column)}${row}`;
        const { html, activeRanges } = getFormulaHTML(
          text,
          cellEditing.sheet,
          model.getTabs().map((tab) => tab.name),
          model.getTokens,
        );
        return {
          ...state,
          cellEditing: {
            ...cellEditing,
            text,
            base: text,
            html,
            activeRanges,
            cursorStart: text.length,
            cursorEnd: text.length,
          },
        };
      }

      case WorkbookActionType.POINTER_MOVE_TO_CELL: {
        const { cell, canvasRef, worksheetRef } = action.payload;
        if (state.cellEditing) {
          return state.reducer(state, {
            type: WorkbookActionType.EDIT_POINTER_MOVE,
            payload: { cell },
          });
        }
        return state.reducer(state, {
          type: WorkbookActionType.EXPAND_AREA_SELECTED_POINTER,
          payload: { cell, canvasRef, worksheetRef },
        });
      }

      case WorkbookActionType.EDIT_POINTER_MOVE: {
        const { model, cellEditing } = state;
        if (!cellEditing) {
          return state;
        }
        const { cell } = action.payload;
        const { base, sheet } = cellEditing;
        let { text } = cellEditing;
        // FIXME: This assumes that cellReference.row > base.row && cellReference.column > base.column
        const cellReference = `${columnNameFromNumber(cell.column)}${cell.row}`;
        if (!base.endsWith(cellReference)) {
          text = `${base}:${cellReference}`;
        }
        const { html, activeRanges } = getFormulaHTML(
          text,
          sheet,
          model.getTabs().map((tab) => tab.name),
          model.getTokens,
        );
        return {
          ...state,
          cellEditing: {
            ...cellEditing,
            text,
            html,
            activeRanges,
            cursorStart: text.length,
            cursorEnd: text.length,
          },
        };
      }

      case WorkbookActionType.EXPAND_AREA_SELECTED_POINTER: {
        const { cell, canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }
        const { row, column } = cell;
        // Magic scrolling by Nico & Mateusz
        const { width, height } = worksheet.getBoundingClientRect();
        const [x, y] = canvas.getCoordinatesByCell(row, column);
        const [x1, y1] = canvas.getCoordinatesByCell(row + 1, column + 1);
        const { left: canvasLeft, top: canvasTop } = canvas.getScrollPosition();
        let border = Border.Right;
        let { left, top } = state.scrollPosition;
        if (x < headerColumnWidth) {
          border = Border.Left;
          left = canvasLeft - headerColumnWidth + x;
        } else if (x1 > width - 20) {
          border = Border.Right;
        }
        if (y < headerRowHeight) {
          border = Border.Top;
          top = canvasTop - headerRowHeight + y;
        } else if (y1 > height - 20) {
          border = Border.Bottom;
        }
        const { selectedCell } = state;
        const area = {
          rowStart: Math.min(selectedCell.row, row),
          rowEnd: Math.max(selectedCell.row, row),
          columnStart: Math.min(selectedCell.column, column),
          columnEnd: Math.max(selectedCell.column, column),
        };
        // If there are frozen rows or columns snap to origin if we cross boundaries
        const frozenRows = canvas.workbook.getFrozenRowsCount();
        const frozenColumns = canvas.workbook.getFrozenColumnsCount();
        if (area.rowStart <= frozenRows && area.rowEnd > frozenRows) {
          top = 0;
        }
        if (area.columnStart <= frozenColumns && area.columnEnd > frozenColumns) {
          left = 0;
        }
        return state.reducer(
          {
            ...state,
            scrollPosition: { left, top },
          },
          {
            type: WorkbookActionType.AREA_SELECTED,
            payload: {
              area,
              border,
              canvasRef,
              worksheetRef,
            },
          },
        );
      }

      case WorkbookActionType.SET_AREA_SELECTED: {
        const { area, border, canvasRef, worksheetRef } = action.payload;
        return state.reducer(
          {
            ...state,
            selectedCell: { column: area.columnStart, row: area.rowStart },
            selectedArea: area,
          },
          {
            type: WorkbookActionType.AREA_SELECTED,
            payload: {
              area,
              border,
              canvasRef,
              worksheetRef,
            },
          },
        );
      }

      case WorkbookActionType.CELL_SELECTED: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }

        const { cell } = action.payload;
        const { row, column } = cell;
        if (row > 0 && row <= canvas.lastRow && column > 0 && column <= canvas.lastColumn) {
          // Magic scrolling by Nico
          const { width, height } = worksheet.getBoundingClientRect();
          const [x, y] = canvas.getCoordinatesByCell(row, column);
          const [x1, y1] = canvas.getCoordinatesByCell(row + 1, column + 1);
          const { left: canvasLeft, top: canvasTop } = canvas.getScrollPosition();
          let { left, top } = state.scrollPosition;

          if (x < headerColumnWidth) {
            left = canvasLeft - headerColumnWidth + x;
          } else if (x1 > width - 20) {
            left = canvasLeft + x1 - width + 20;
          }
          if (y < headerRowHeight) {
            top = canvasTop - headerRowHeight + y;
          } else if (y1 > height - 20) {
            top = canvasTop + y1 - height + 20;
          }

          return state.reducer(
            {
              ...state,
              selectedCell: { row, column },
              selectedArea: { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
              scrollPosition: { left, top },
              cellEditing: null,
            },
            { type: WorkbookActionType.SELECTED_CELL_HAS_CHANGED },
          );
        }
        return state;
      }

      case WorkbookActionType.NAVIGATION_TO_EDGE: {
        const { model, selectedSheet, selectedCell } = state;
        const { key, canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        if (!canvas) {
          return state;
        }
        const newSelectedCell = model.getNavigationEdge(
          key,
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
          canvas.lastRow,
          canvas.lastColumn,
        );
        return state.reducer(state, {
          type: WorkbookActionType.CELL_SELECTED,
          payload: { cell: newSelectedCell, canvasRef, worksheetRef },
        });
      }

      case WorkbookActionType.EXTEND_TO_CELL: {
        const { selectedArea } = state;
        const { cell, canvasRef, worksheetRef } = action.payload;
        const area = getExpandToArea(selectedArea, cell);
        return state.reducer(state, {
          type: WorkbookActionType.EXTEND_TO_SELECTED,
          payload: {
            area,
            canvasRef,
            worksheetRef,
          },
        });
      }

      case WorkbookActionType.EXTEND_TO_SELECTED: {
        const { canvasRef, worksheetRef } = action.payload;
        const canvas = canvasRef.current;
        const worksheet = worksheetRef.current;
        if (!canvas || !worksheet) {
          return state;
        }

        const { area } = action.payload;
        if (!area) {
          return {
            ...state,
            extendToArea: null,
          };
        }

        const { rowStart, rowEnd, columnStart, columnEnd } = area;
        let { left, top } = state.scrollPosition;

        // If there are frozen rows snap to origin if we cross boundaries
        const frozenRows = canvas.workbook.getFrozenRowsCount();
        const frozenColumns = canvas.workbook.getFrozenColumnsCount();
        // NOTE: rowStart in this case is the first row of the "extend-to" action,
        // so there is at least one row on top
        if (rowStart - 1 <= frozenRows && rowEnd > frozenRows) {
          top = 0;
        }
        if (columnStart - 1 <= frozenColumns && columnEnd > frozenColumns) {
          left = 0;
        }

        if (
          rowStart > 0 &&
          columnStart > 0 &&
          rowEnd <= canvas.lastRow &&
          columnEnd <= canvas.lastColumn
        ) {
          return {
            ...state,
            scrollPosition: { left, top },
            extendToArea: { ...area },
          };
        }
        return state;
      }

      case WorkbookActionType.EXTEND_TO_END: {
        const { model, extendToArea, selectedSheet, selectedArea } = state;
        if (!extendToArea) {
          return state;
        }

        const { rowStart, rowEnd, columnStart, columnEnd } = extendToArea;
        model.extendTo(selectedSheet, selectedArea, { rowStart, rowEnd, columnStart, columnEnd });
        return state.reducer(
          {
            ...state,
            selectedArea: mergedAreas({ rowStart, rowEnd, columnStart, columnEnd }, selectedArea),
            extendToArea: null,
          },
          { type: WorkbookActionType.WORKBOOK_HAS_CHANGED },
        );
      }

      case WorkbookActionType.UNDO: {
        const { model } = state;
        if (model.readOnly) {
          return state;
        }
        model.undo();
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.REDO: {
        const { model } = state;
        if (model.readOnly) {
          return state;
        }
        model.redo();
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.DELETE_CELLS: {
        const { model, selectedArea, selectedSheet } = state;
        if (model.readOnly) {
          return state;
        }
        model.deleteCells(selectedSheet, selectedArea);
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_BOLD: {
        const { model, selectedSheet, selectedArea, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        const b = !selectedStyle.font.b;
        model.setCellsStyle(selectedSheet, selectedArea, (style: CellStyle): CellStyle => {
          const font = {
            ...style.font,
            b,
          };
          return {
            ...style,
            font,
          };
        });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_ITALIC: {
        const { model, selectedSheet, selectedArea, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        model.setCellsStyle(selectedSheet, selectedArea, (style: CellStyle): CellStyle => {
          const font = {
            ...style.font,
            i: !selectedStyle.font.i,
          };
          return {
            ...style,
            font,
          };
        });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_UNDERLINE: {
        const { model, selectedSheet, selectedArea, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        const u = !selectedStyle.font.u;
        model.setCellsStyle(selectedSheet, selectedArea, (style: CellStyle): CellStyle => {
          const font = {
            ...style.font,
            u,
          };
          return {
            ...style,
            font,
          };
        });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_STRIKE: {
        const { model, selectedSheet, selectedArea, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        const strike = !selectedStyle.font.strike;
        model.setCellsStyle(selectedSheet, selectedArea, (style: CellStyle): CellStyle => {
          const font = {
            ...style.font,
            strike,
          };
          return {
            ...style,
            font,
          };
        });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_ALIGN_LEFT: {
        const { model, selectedArea, selectedSheet, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        const horizontalAlignment =
          selectedStyle.horizontal_alignment === 'left' ? 'default' : 'left';
        model.setCellsStyle(
          selectedSheet,
          selectedArea,
          (style: CellStyle): CellStyle => ({
            ...style,
            horizontal_alignment: horizontalAlignment,
          }),
        );

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_ALIGN_CENTER: {
        const { model, selectedArea, selectedSheet, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        const horizontalAlignment =
          selectedStyle.horizontal_alignment === 'center' ? 'default' : 'center';
        model.setCellsStyle(
          selectedSheet,
          selectedArea,
          (style: CellStyle): CellStyle => ({
            ...style,
            horizontal_alignment: horizontalAlignment,
          }),
        );

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TOGGLE_ALIGN_RIGHT: {
        const { model, selectedArea, selectedSheet, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }
        const selectedStyle = model.getCellStyle(
          selectedSheet,
          selectedCell.row,
          selectedCell.column,
        );
        const horizontalAlignment =
          selectedStyle.horizontal_alignment === 'right' ? 'default' : 'right';
        model.setCellsStyle(
          selectedSheet,
          selectedArea,
          (style: CellStyle): CellStyle => ({
            ...style,
            horizontal_alignment: horizontalAlignment,
          }),
        );

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.TEXT_COLOR_PICKED: {
        const { model, selectedSheet, selectedArea } = state;
        if (model.readOnly) {
          return state;
        }
        model.setCellsStyle(selectedSheet, selectedArea, (style: CellStyle): CellStyle => {
          const font = {
            ...style.font,
            color: {
              RGB: action.payload.color,
            },
          };
          return {
            ...style,
            font,
          };
        });
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.SHEET_COLOR_CHANGED: {
        const { model, selectedSheet } = state;
        if (model.readOnly) {
          return state;
        }
        model.setSheetColor(selectedSheet, action.payload.color);
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.SHEET_DELETED: {
        const { model, selectedSheet } = state;
        if (model.readOnly) {
          return state;
        }
        model.deleteSheet(selectedSheet);
        return state.reducer(
          { ...state, selectedSheet: 0 },
          { type: WorkbookActionType.WORKBOOK_HAS_CHANGED },
        );
      }

      case WorkbookActionType.SHEET_RENAMED: {
        const { model, selectedSheet } = state;
        if (model.readOnly) {
          return state;
        }
        model.renameSheet(selectedSheet, action.payload.newName);
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.ADD_BLANK_SHEET: {
        const { model } = state;
        if (model.readOnly) {
          return state;
        }
        model.addBlankSheet();
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.FILL_COLOR_PICKED: {
        const { model, selectedSheet, selectedArea } = state;
        if (model.readOnly) {
          return state;
        }

        model.setCellsStyle(selectedSheet, selectedArea, (style: CellStyle): CellStyle => {
          const fill = {
            ...style.fill,
            fg_color: {
              RGB: action.payload.color,
            },
          };
          return {
            ...style,
            fill,
          };
        });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.NUMBER_FORMAT_PICKED: {
        const { model, selectedSheet, selectedArea } = state;
        if (model.readOnly) {
          return state;
        }
        model.setCellsStyle(
          selectedSheet,
          selectedArea,
          (style: CellStyle): CellStyle => ({
            ...style,
            num_fmt: action.payload.numFmt,
          }),
        );

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.CHANGE_COLUMN_WIDTH: {
        const { model } = state;
        const { sheet, column, width, onResize } = action.payload;
        // Minimum width is 2px
        const newColumnWidth = Math.max(2, width);
        const oldColumnWidth = model.getColumnWidth(sheet, column);

        model.setColumnWidth(sheet, column, newColumnWidth);

        onResize({ deltaWidth: newColumnWidth - oldColumnWidth, deltaHeight: 0 });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.CHANGE_ROW_HEIGHT: {
        const { model } = state;
        const { sheet, row, height, onResize } = action.payload;
        // Minimum height is 2px
        const newRowHeight = Math.max(2, height);
        const oldRowHeight = model.getRowHeight(sheet, row);

        model.setRowHeight(sheet, row, newRowHeight);

        onResize({ deltaHeight: newRowHeight - oldRowHeight, deltaWidth: 0 });

        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.DELETE_ROW: {
        const { model } = state;
        const { sheet, row } = action.payload;
        model.deleteRow(sheet, row);
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.INSERT_ROW: {
        const { model } = state;
        const { sheet, row } = action.payload;
        model.insertRow(sheet, row);
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      /**
       * The clipboard allows us to attach different values to different mime types.
       * When copying we return two things: A TSV string (tab separated values).
       * And a an string representing the area we are copying.
       * We attach the tsv string to "text/plain" useful to paste to a different application
       * We attach the area to 'application/json' useful to paste within the application.
       *
       * FIXME: This second part is cheesy and will produce unexpected results:
       *      1. User copies an area (call it's contents area1)
       *      2. User modifies the copied area (call it's contents area2)
       *      3. User paste content to a different place within the application
       *
       * To the user surprise area2 will be pasted. The fix for this ius ro actually return a json with the actual content.
       */
      case WorkbookActionType.COPY: {
        const { model, selectedSheet, selectedArea } = state;
        const { tsv, area, sheetData } = model.copy({ sheet: selectedSheet, ...selectedArea });
        const { event } = action.payload;
        let clipboardId = sessionStorage.getItem(CLIPBOARD_ID_SESSION_STORAGE_KEY);
        if (!clipboardId) {
          clipboardId = getNewClipboardId();
          sessionStorage.setItem(CLIPBOARD_ID_SESSION_STORAGE_KEY, clipboardId);
        }
        event.clipboardData.setData('text/plain', tsv);
        event.clipboardData.setData(
          'application/json',
          JSON.stringify({ type: 'copy', area, sheetData, clipboardId }),
        );
        return state;
      }

      case WorkbookActionType.PASTE: {
        const { model, selectedSheet, selectedArea, selectedCell } = state;
        if (model.readOnly) {
          return state;
        }

        const { format, data, event } = action.payload;
        if (format === 'application/json') {
          // We are copying from within the application
          const targetArea = {
            sheet: selectedSheet,
            ...selectedArea,
          };

          try {
            const source = JSON.parse(data);
            const clipboardId = sessionStorage.getItem(CLIPBOARD_ID_SESSION_STORAGE_KEY);
            let sourceType = source.type;
            if (clipboardId !== source.clipboardId) {
              sourceType = 'copy';
            }
            model.paste(source.area, targetArea, source.sheetData, sourceType);
            if (sourceType === 'cut') {
              event.clipboardData.clearData();
            }
          } catch {
            // Trying to paste incorrect JSON
            // FIXME: We should validate the JSON and not try/catch
            // If JSON is invalid We should try to paste 'text/plain' content if it exist
          }
        } else if (format === 'text/plain') {
          model.pasteText(selectedSheet, selectedCell, data);
        } else {
          // NOT IMPLEMENTED
        }
        return state.reducer(state, { type: WorkbookActionType.WORKBOOK_HAS_CHANGED });
      }

      case WorkbookActionType.CUT: {
        const { model, selectedSheet, selectedArea } = state;
        const { tsv, area, sheetData } = model.copy({ sheet: selectedSheet, ...selectedArea });
        const { event } = action.payload;
        let clipboardId = sessionStorage.getItem(CLIPBOARD_ID_SESSION_STORAGE_KEY);
        if (!clipboardId) {
          clipboardId = getNewClipboardId();
          sessionStorage.setItem(CLIPBOARD_ID_SESSION_STORAGE_KEY, clipboardId);
        }
        event.clipboardData.setData('text/plain', tsv);
        event.clipboardData.setData(
          'application/json',
          JSON.stringify({ type: 'cut', area, sheetData, clipboardId }),
        );
        return state;
      }

      default:
        throw new Error(`Unsupported action called ${action}`);
    }
  }
};

const useWorkbookReducer = (
  reducer = defaultWorkbookReducer,
): [WorkbookState, Dispatch<Action>] => {
  const [state, dispatch] = useReducer(reducer, getInitialWorkbookState(reducer));

  useEffect(() => {
    dispatch({ type: WorkbookActionType.CHANGE_REDUCER, payload: { reducer } });
  }, [reducer]);

  return [state, dispatch];
};

export default useWorkbookReducer;
