import { useReducer, Dispatch, useRef } from 'react';
import {
  Border,
  StateSettings,
  mergedAreas,
  FocusType,
  columnNameFromNumber,
  getExpandToArea,
  quoteSheetName,
} from 'src/components/workbook/util';
import { WorkbookActionType, Action, WorkbookReducer, WorkbookState } from './common';
import { headerRowHeight, headerColumnWidth } from '../canvas';
import Model from '../model';

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

export const getInitialWorkbookState = (
  modelRef: React.RefObject<Model | null>,
  reducer: WorkbookReducer,
): WorkbookState => ({
  modelRef,
  selectedSheet: 0,
  sheetStates: {},
  cellEditing: null,
  scrollPosition: { ...defaultSheetState.scrollPosition },
  selectedCell: { ...defaultSheetState.selectedCell },
  selectedArea: { ...defaultSheetState.selectedArea },
  extendToArea: null,
  reducer,
});

export const defaultWorkbookReducer: WorkbookReducer = (state, action): WorkbookState => {
  if (state.modelRef.current === null) {
    return state;
  }
  switch (action.type) {
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
      const { sheet, row, column } = cellEditing;

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

      return { ...newState, cellEditing: null };
    }

    case WorkbookActionType.EDIT_ESCAPE: {
      const { cellEditing } = state;
      if (!cellEditing) {
        return state;
      }
      const { sheet, row, column } = cellEditing;
      return {
        ...state,
        cellEditing: null,
        selectedSheet: sheet,
        selectedCell: { row, column },
        selectedArea: { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
      };
    }

    case WorkbookActionType.EDIT_KEY_PRESS_START: {
      const { selectedSheet, selectedCell } = state;
      const { initText } = action.payload;

      return {
        ...state,
        cellEditing: {
          sheet: selectedSheet,
          row: selectedCell.row,
          column: selectedCell.column,
          text: initText,
          base: initText,
          mode: 'init',
          focus: FocusType.Cell,
        },
      };
    }

    case WorkbookActionType.EDIT_CELL_EDITOR_START: {
      const { selectedSheet, selectedCell, cellEditing } = state;
      const model = state.modelRef.current;
      if (cellEditing) {
        // Ignore if already editing
        return state;
      }
      const { row, column } = selectedCell;
      let text = model.getFormulaOrValue(selectedSheet, row, column);
      if (!action.payload.ignoreQuotePrefix && model.isQuotePrefix(selectedSheet, row, column)) {
        text = `'${text}`;
      }
      return {
        ...state,
        cellEditing: {
          text,
          base: text,
          focus: FocusType.Cell,
          sheet: selectedSheet,
          row,
          column,
          mode: 'edit',
        },
      };
    }

    case WorkbookActionType.EDIT_FORMULA_BAR_EDITOR_START: {
      const { cellEditing } = state;
      const model = state.modelRef.current;
      if (cellEditing) {
        return state;
      }

      // Start completely new edit
      const { selectedCell, selectedSheet } = state;
      const { row, column } = selectedCell;
      let text = model.getFormulaOrValue(selectedSheet, row, column);
      if (model.isQuotePrefix(selectedSheet, row, column)) {
        text = `'${text}`;
      }
      return {
        ...state,
        cellEditing: {
          text,
          base: text,
          focus: FocusType.FormulaBar,
          sheet: selectedSheet,
          row,
          column,
          mode: 'edit',
        },
      };
    }

    case WorkbookActionType.SELECT_SHEET: {
      const { sheetStates } = state;
      const model = state.modelRef.current;
      const tabs = model.getTabs();
      if (tabs[state.selectedSheet]) {
        sheetStates[tabs[state.selectedSheet].sheet_id] = {
          scrollPosition: { ...state.scrollPosition },
          selectedCell: { ...state.selectedCell },
          selectedArea: { ...state.selectedArea },
          extendToArea: state.extendToArea && { ...state.extendToArea },
        };
      }
      let { sheet } = action.payload;

      // In case the sheet does not exist anymore (i.e. has been deleted)
      if (tabs.length <= sheet) {
        sheet = 0;
      }

      const sheetId = tabs[sheet].sheet_id;
      const sheetState = sheetId in sheetStates ? sheetStates[sheetId] : defaultSheetState;

      return {
        ...state,
        selectedSheet: sheet,
        sheetStates: { ...sheetStates },
        selectedCell: { ...sheetState.selectedCell },
        selectedArea: { ...sheetState.selectedArea },
        scrollPosition: { ...sheetState.scrollPosition },
        extendToArea: sheetState.extendToArea && { ...sheetState.extendToArea },
        cellEditing: state.cellEditing && { ...state.cellEditing, focus: FocusType.FormulaBar },
      };
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

      return {
        ...state,
        selectedCell: { ...state.selectedCell, row: newRow },
        selectedArea: {
          columnStart: column,
          columnEnd: column,
          rowStart: newRow,
          rowEnd: newRow,
        },
        scrollPosition: { ...state.scrollPosition, top: scrollTop },
      };
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

      return {
        ...state,
        selectedCell: { ...state.selectedCell, column: newColumn },
        selectedArea: {
          rowStart: row,
          rowEnd: row,
          columnStart: newColumn,
          columnEnd: newColumn,
        },
        scrollPosition: { ...state.scrollPosition, left: scrollLeft },
      };
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

      return {
        ...state,
        selectedCell: { ...state.selectedCell, column: newColumn },
        selectedArea: {
          rowStart: row,
          rowEnd: row,
          columnStart: newColumn,
          columnEnd: newColumn,
        },
        scrollPosition: { ...state.scrollPosition, left: scrollLeft },
      };
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
        const newCellBottom = canvas.getCoordinatesByCell(newRow + 1, state.selectedCell.column)[1];
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

      return {
        ...state,
        selectedCell: { ...state.selectedCell, row: newRow },
        selectedArea: {
          rowStart: newRow,
          rowEnd: newRow,
          columnStart: column,
          columnEnd: column,
        },
        scrollPosition: { ...state.scrollPosition, top: scrollTop },
      };
    }

    case WorkbookActionType.KEY_END: {
      const { canvasRef, worksheetRef } = action.payload;
      const canvas = canvasRef.current;
      const worksheet = worksheetRef.current;
      if (!canvas || !worksheet) {
        return state;
      }
      const { row } = state.selectedCell;
      return {
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
      };
    }

    case WorkbookActionType.KEY_HOME: {
      const { canvasRef, worksheetRef } = action.payload;
      const canvas = canvasRef.current;
      const worksheet = worksheetRef.current;
      if (!canvas || !worksheet) {
        return state;
      }
      const { row } = state.selectedCell;
      return {
        ...state,
        selectedCell: { ...state.selectedCell, column: 1 },
        selectedArea: { columnStart: 1, columnEnd: 1, rowStart: row, rowEnd: row },
        scrollPosition: { ...state.scrollPosition, left: 0 },
      };
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

      return {
        ...state,
        selectedCell: { ...state.selectedCell, row: newRow },
        selectedArea: {
          rowStart: newRow,
          rowEnd: newRow,
          columnStart: column,
          columnEnd: column,
        },
        scrollPosition: { ...state.scrollPosition, top: scrollTop },
      };
    }

    case WorkbookActionType.PAGE_UP: {
      const { canvasRef, worksheetRef } = action.payload;
      const model = state.modelRef.current;
      const canvas = canvasRef.current;
      const worksheet = worksheetRef.current;
      if (!canvas || !worksheet) {
        return state;
      }

      // Compute scroll - magic by Nico
      const { topLeftCell } = canvas.getVisibleCells();
      const delta = state.selectedCell.row - topLeftCell.row;
      const topCellRowHeight = model.getRowHeight(state.selectedSheet, topLeftCell.row);
      const deltaHeight = canvas.height - headerRowHeight - topCellRowHeight;
      const scrollTop = Math.max(0, canvas.getScrollPosition().top - deltaHeight);

      const newRow = canvas.getBoundedRow(scrollTop).row + delta;
      const { column } = state.selectedCell;

      return {
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
      };
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
      const { cell, canvasRef, worksheetRef } = action.payload;
      if (cellEditing) {
        const { sheet } = cellEditing;
        const newRow = cellEditing.row;
        const newColumn = cellEditing.column;
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

        return { ...newState, cellEditing: null };
      }
      return state.reducer(state, {
        type: WorkbookActionType.CELL_SELECTED,
        payload: { cell, canvasRef, worksheetRef },
      });
    }

    case WorkbookActionType.EDIT_POINTER_DOWN: {
      const { cellEditing, selectedSheet } = state;
      const model = state.modelRef.current;
      const tabs = model.getTabs();
      if (!cellEditing) {
        return state;
      }
      const { cell, currentValue } = action.payload;
      const { row, column } = cell;
      const prefix =
        cellEditing.sheet === selectedSheet ? '' : `${quoteSheetName(tabs[selectedSheet].name)}!`;
      const text = `${currentValue}${prefix}${columnNameFromNumber(column)}${row}`;
      return {
        ...state,
        cellEditing: {
          ...cellEditing,
          text,
          base: text,
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
      const { cellEditing } = state;
      if (!cellEditing) {
        return state;
      }
      const { cell } = action.payload;
      const { base } = cellEditing;
      let { text } = cellEditing;
      // FIXME: This assumes that cellReference.row > base.row && cellReference.column > base.column
      const cellReference = `${columnNameFromNumber(cell.column)}${cell.row}`;
      if (!base.endsWith(cellReference)) {
        text = `${base}:${cellReference}`;
      }
      return {
        ...state,
        cellEditing: {
          ...cellEditing,
          text,
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

    case WorkbookActionType.CELL_SELECTED: {
      const { canvasRef, worksheetRef } = action.payload;
      const canvas = canvasRef.current;
      const worksheet = worksheetRef.current;
      if (!canvas || !worksheet) {
        return state;
      }

      const { cell } = action.payload;
      let { row, column } = cell;
      row = Math.max(0, Math.min(row, canvas.lastRow));
      column = Math.max(0, Math.min(column, canvas.lastColumn));

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

      return {
        ...state,
        selectedCell: { row, column },
        selectedArea: { rowStart: row, rowEnd: row, columnStart: column, columnEnd: column },
        scrollPosition: { left, top },
        cellEditing: null,
      };
    }

    case WorkbookActionType.NAVIGATION_TO_EDGE: {
      const { selectedSheet, selectedCell } = state;
      const model = state.modelRef.current;
      const { key, canvasRef, worksheetRef } = action.payload;
      const cell = model.getNavigationEdge(
        key,
        selectedSheet,
        selectedCell.row,
        selectedCell.column,
      );
      return state.reducer(state, {
        type: WorkbookActionType.CELL_SELECTED,
        payload: {
          cell,
          canvasRef,
          worksheetRef,
        },
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
      const { extendToArea, selectedArea } = state;
      if (!extendToArea) {
        return state;
      }
      return {
        ...state,
        selectedArea: mergedAreas(extendToArea, selectedArea),
        extendToArea: null,
      };
    }

    default:
      throw new Error(`Unsupported action called ${action}`);
  }
};

const useWorkbookReducer = (
  model: Model | null,
  reducer = defaultWorkbookReducer,
): [WorkbookState, Dispatch<Action>] => {
  const modelRef = useRef(model);
  modelRef.current = model;
  const [state, dispatch] = useReducer(reducer, getInitialWorkbookState(modelRef, reducer));
  return [state, dispatch];
};

export default useWorkbookReducer;
