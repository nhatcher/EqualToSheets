import React, { FunctionComponent, useEffect, useRef, useState, useCallback } from 'react';
import styled from 'styled-components';
import * as Menu from 'src/components/uiKit/menu';
import WorksheetCanvas from './canvas';
import { Cell } from './util';
import usePointer from './usePointer';
import useResize from './useResize';
import Editor from './editor';
import { outlineBackgroundColor, outlineColor } from './constants';
import { defaultSheetState } from './useWorkbookReducer';
import useScrollSync from './useScrollSync';
import { RowContextMenuContent, ColumnContextMenuContent } from './contextMenuContent';
import { isInReferenceMode } from './formulas';
import { useWorkbookContext } from './workbookContext';

export type WorkbookElements = {
  cellEditor: HTMLDivElement | null;
};

const Workbook: FunctionComponent<{
  className?: string;
}> = (properties) => {
  const {
    model,
    focusWorkbook,
    requestRender,
    requestRenderId,
    editorState,
    editorActions,
    worksheetCanvas,
    worksheetElement,
    cellInput,
    formulaBarInput,
    lastRow,
    lastColumn,
  } = useWorkbookContext();
  const { selectedSheet, selectedCell, selectedArea, cellEditing, extendToArea, scrollPosition } =
    editorState;

  // References to DOM elements
  const canvasElement = useRef<HTMLCanvasElement>(null);
  const scrollElement = useRef<HTMLDivElement>(null);
  const spacerElement = useRef<HTMLDivElement>(null);
  const cellOutline = useRef<HTMLDivElement>(null);
  const areaOutline = useRef<HTMLDivElement>(null);
  const cellOutlineHandle = useRef<HTMLDivElement>(null);
  const extendToOutline = useRef<HTMLDivElement>(null);
  const columnResizeGuide = useRef<HTMLDivElement>(null);
  const rowResizeGuide = useRef<HTMLDivElement>(null);
  const rowContextMenuAnchorElement = useRef<HTMLDivElement>(null);
  const columnContextMenuAnchorElement = useRef<HTMLDivElement>(null);
  const columnHeaders = useRef<HTMLDivElement>(null);

  const resizeSpacer = useCallback((options: { deltaWidth: number; deltaHeight: number }): void => {
    if (spacerElement.current) {
      const spacerRect = spacerElement.current.getBoundingClientRect();

      const newSpacerWidth = spacerRect.width + options.deltaWidth;
      const newSpacerHeight = spacerRect.height + options.deltaHeight;

      spacerElement.current.style.width = `${newSpacerWidth}px`;
      spacerElement.current.style.height = `${newSpacerHeight}px`;
    }
  }, []);

  const onColumnWidthChange = useCallback(
    (sheet: number, column: number, width: number): void => {
      if (!model) {
        return;
      }
      // Minimum width is 2px
      const newColumnWidth = Math.max(2, width);
      const oldColumnWidth = model.getColumnWidth(sheet, column);

      model.setColumnWidth(sheet, column, newColumnWidth);

      resizeSpacer({ deltaWidth: newColumnWidth - oldColumnWidth, deltaHeight: 0 });
    },
    [model, resizeSpacer],
  );

  const onRowHeightChange = useCallback(
    (sheet: number, row: number, height: number): void => {
      if (!model) {
        return;
      }
      // Minimum height is 2px
      const newRowHeight = Math.max(2, height);
      const oldRowHeight = model.getRowHeight(sheet, row);

      model.setRowHeight(sheet, row, newRowHeight);

      resizeSpacer({ deltaHeight: newRowHeight - oldRowHeight, deltaWidth: 0 });
    },
    [model, resizeSpacer],
  );

  const onPointerDownAtCell = (cell: Cell, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    if (cellEditing && cellInput.current && formulaBarInput.current) {
      const { value } = cellInput.current;
      const selection = window.getSelection();
      if (
        selection?.anchorNode?.lastChild &&
        (selection.anchorNode.lastChild === cellInput.current ||
          selection.anchorNode.lastChild === formulaBarInput.current)
      ) {
        const cursor =
          (selection.anchorNode.lastChild as HTMLInputElement).selectionEnd ?? value.length;
        if (isInReferenceMode(value, cursor)) {
          editorActions.onEditPointerDown(cell, value);
          return;
        }
      }
      model?.setCellValue(cellEditing.sheet, cellEditing.row, cellEditing.column, value);
    }
    editorActions.onPointerDownAtCell(cell);
    focusWorkbook();
  };

  const { getCanvasSize, onResize } = useResize(worksheetElement, worksheetCanvas);
  const { onScroll } = useScrollSync(
    scrollElement,
    scrollPosition,
    editorActions.setScrollPosition,
  );

  const [isRowContextMenuOpen, setIsRowContextMenuOpen] = useState(false);
  const [rowContextMenu, setRowContextMenu] = useState(0);
  const onRowContextMenu = useCallback(
    (row: number): void => {
      if (!model) {
        return;
      }
      setIsRowContextMenuOpen(true);
      setRowContextMenu(row);
    },
    [model],
  );

  const onInsertRow = (row: number): void => {
    model?.insertRow(selectedSheet, row);
    setIsRowContextMenuOpen(false);
  };

  const onDeleteRow = (row: number): void => {
    model?.deleteRow(selectedSheet, row);
    setIsRowContextMenuOpen(false);
  };

  const [isColumnContextMenuOpen, setIsColumnContextMenuOpen] = useState(false);
  const [columnContextMenu, setColumnContextMenu] = useState(0);
  const onColumnContextMenu = useCallback(
    (column: number): void => {
      if (!model) {
        return;
      }
      setIsColumnContextMenuOpen(true);
      setColumnContextMenu(column);
    },
    [model],
  );

  const onInsertColumn = (column: number): void => {
    model?.insertColumn(selectedSheet, column);
    setIsColumnContextMenuOpen(false);
  };

  const onDeleteColumn = (column: number): void => {
    model?.deleteColumn(selectedSheet, column);
    setIsColumnContextMenuOpen(false);
  };

  const { onPointerMove, onPointerDown, onPointerHandleDown, onPointerUp, onContextMenu } =
    usePointer({
      onAreaSelected: editorActions.onAreaSelected,
      onPointerDownAtCell,
      onPointerMoveToCell: editorActions.onPointerMoveToCell,
      onExtendToCell: editorActions.onExtendToCell,
      onExtendToEnd: editorActions.onExtendToEnd,
      canvasElement,
      worksheetElement,
      worksheetCanvas,
      rowContextMenuAnchorElement,
      columnContextMenuAnchorElement,
      onRowContextMenu,
      onColumnContextMenu,
    });

  // Init canvas
  useEffect(() => {
    const outline = cellOutline.current;
    const handle = cellOutlineHandle.current;
    const area = areaOutline.current;
    const canvas = canvasElement.current;
    const worksheetWrapper = worksheetElement.current;
    const extendTo = extendToOutline.current;
    const columnGuide = columnResizeGuide.current;
    const rowGuide = rowResizeGuide.current;
    const columnHeadersDiv = columnHeaders.current;
    // Silence the linter
    if (
      !model ||
      !outline ||
      !handle ||
      !area ||
      !canvas ||
      !worksheetWrapper ||
      !extendTo ||
      !rowGuide ||
      !columnGuide ||
      !columnHeadersDiv
    ) {
      return (): void => {};
    }

    const canvasSize = getCanvasSize();
    if (!canvasSize) {
      throw new Error('Could not find the canvas element size.');
    }

    const canvasSettings = {
      model,
      selectedSheet: 0,
      width: canvasSize.width,
      height: canvasSize.height,
      elements: {
        cellOutline: outline,
        cellOutlineHandle: handle,
        areaOutline: area,
        extendToOutline: extendTo,
        columnGuide,
        rowGuide,
        canvas,
        columnHeaders: columnHeadersDiv,
      },
      state: {
        ...defaultSheetState,
      },
      cellEditing: null,
      onColumnWidthChanges: onColumnWidthChange,
      onRowHeightChanges: onRowHeightChange,
      lastRow,
      lastColumn,
    };

    worksheetCanvas.current = new WorksheetCanvas(canvasSettings);

    const [sheetWidth, sheetHeight] = worksheetCanvas.current.getSheetDimensions();
    if (spacerElement.current) {
      spacerElement.current.style.height = `${sheetHeight}px`;
      spacerElement.current.style.width = `${sheetWidth}px`;
    }

    const resizeObserver = new ResizeObserver(onResize);
    resizeObserver.observe(worksheetElement.current);
    requestRender();

    return (): void => {
      worksheetCanvas.current = null;
      resizeObserver.disconnect();
    };
  }, [
    requestRender,
    getCanvasSize,
    model,
    onColumnWidthChange,
    onResize,
    onRowHeightChange,
    worksheetElement,
    worksheetCanvas,
    lastRow,
    lastColumn,
  ]);

  // Sync canvas with sheet state
  useEffect(() => {
    const state = {
      scrollPosition,
      selectedCell,
      selectedArea,
      extendToArea,
    };
    worksheetCanvas.current?.setSelectedSheet(selectedSheet);
    worksheetCanvas.current?.setState(state);
    requestRender();
  }, [
    extendToArea,
    requestRender,
    scrollPosition,
    selectedArea,
    selectedCell,
    selectedSheet,
    worksheetCanvas,
  ]);

  // Sync canvas with editing state
  useEffect(() => {
    worksheetCanvas.current?.setCellEditing(cellEditing);
    requestRender();
  }, [cellEditing, requestRender, worksheetCanvas]);

  useEffect(() => {
    worksheetCanvas.current?.renderSheet();
  }, [requestRenderId, worksheetCanvas]);

  const stopPropagationIfEditing = (event: any) => {
    if (cellEditing) {
      event.stopPropagation();
    }
  };

  return (
    <Worksheet
      className={properties.className}
      $hidden={!model}
      onScroll={onScroll}
      ref={scrollElement}
    >
      <Spacer ref={spacerElement} />
      <SheetCanvasWrapper
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onDoubleClick={(event) => {
          editorActions.onCellEditStart();
          event.stopPropagation();
          event.preventDefault();
        }}
        onContextMenu={onContextMenu}
        ref={worksheetElement}
      >
        <SheetCanvas ref={canvasElement} />
        <CellOutline
          ref={cellOutline}
          onPointerDown={stopPropagationIfEditing}
          onPointerUp={stopPropagationIfEditing}
          onPointerMove={stopPropagationIfEditing}
          onCopy={stopPropagationIfEditing}
          onCut={stopPropagationIfEditing}
          onPaste={stopPropagationIfEditing}
        >
          <Editor
            onReferencesChanged={(references) => {
              if (worksheetCanvas.current) {
                worksheetCanvas.current.activeRanges = references;
                worksheetCanvas.current.renderSheet();
              }
            }}
          />
        </CellOutline>
        <AreaOutline ref={areaOutline} />
        <ExtendToOutline ref={extendToOutline} />
        <CellOutlineHandle ref={cellOutlineHandle} onPointerDown={onPointerHandleDown} />
        <ColumnResizeGuide ref={columnResizeGuide} />
        <RowResizeGuide ref={rowResizeGuide} />
        <ColumnHeaders ref={columnHeaders} />
      </SheetCanvasWrapper>
      <Menu.Root open={isRowContextMenuOpen} onOpenChange={setIsRowContextMenuOpen}>
        <Menu.Trigger asChild>
          <ContextMenuAnchorElement ref={rowContextMenuAnchorElement} />
        </Menu.Trigger>
        <RowContextMenuContent
          isMenuOpen={isRowContextMenuOpen}
          row={rowContextMenu}
          anchorEl={rowContextMenuAnchorElement.current}
          onDeleteRow={onDeleteRow}
          onInsertRow={onInsertRow}
          onClose={(): void => {
            setIsRowContextMenuOpen(false);
          }}
        />
      </Menu.Root>
      <Menu.Root open={isColumnContextMenuOpen} onOpenChange={setIsColumnContextMenuOpen}>
        <Menu.Trigger asChild>
          <ContextMenuAnchorElement ref={columnContextMenuAnchorElement} />
        </Menu.Trigger>
        <ColumnContextMenuContent
          isMenuOpen={isColumnContextMenuOpen}
          column={columnContextMenu}
          anchorEl={columnContextMenuAnchorElement.current}
          onDeleteColumn={onDeleteColumn}
          onInsertColumn={onInsertColumn}
          onClose={(): void => {
            setIsColumnContextMenuOpen(false);
          }}
        />
      </Menu.Root>
    </Worksheet>
  );
};

export default Workbook;

const ContextMenuAnchorElement = styled.div`
  position: absolute;
  width: 0px;
  height: 0px;
  visibility: hidden;
`;

const ColumnResizeGuide = styled.div`
  position: absolute;
  top: 0px;
  display: none;
  height: 100%;
  width: 0px;
  border-left: 1px dashed ${outlineColor};
`;

const ColumnHeaders = styled.div`
  position: absolute;
  left: 0px;
  top: 0px;
  overflow: hidden;
  & .column-header {
    display: inline-block;
    text-align: center;
    overflow: hidden;
    height: 100%;
    user-select: none;
  }
`;

const RowResizeGuide = styled.div`
  position: absolute;
  display: none;
  left: 0px;
  height: 0px;
  width: 100%;
  border-top: 1px dashed ${outlineColor};
`;

const AreaOutline = styled.div`
  box-sizing: border-box;
  position: absolute;
  border: 1px solid ${outlineColor};
  border-radius: 3px;
  background-color: ${outlineBackgroundColor};
`;

const CellOutline = styled.div`
  box-sizing: border-box;
  position: absolute;
  border: 2px solid ${outlineColor};
  border-radius: 1px;
  box-shadow: 0px 1px 4px rgba(0, 0, 0, 0.15);
  word-break: break-word;
  font-size: 13px;
  display: flex;
`;

const CellOutlineHandle = styled.div`
  position: absolute;
  width: 7px;
  height: 7px;
  background: ${outlineColor};
  cursor: crosshair;
  border: 1px solid white;
  border-radius: 50%;
`;

const ExtendToOutline = styled.div`
  box-sizing: border-box;
  position: absolute;
  border: 1px dashed ${outlineColor};
  border-radius: 3px;
`;

const SheetCanvasWrapper = styled.div`
  position: sticky;
  top: 0px;
  left: 0px;
  height: 100%;

  .column-resize-handle {
    position: absolute;
    top: 0px;
    width: 3px;
    opacity: 0;
    background: ${outlineColor};
    border-radius: 5px;
    cursor: col-resize;
  }

  .column-resize-handle:hover {
    opacity: 1;
  }
  .row-resize-handle {
    position: absolute;
    left: 0px;
    height: 3px;
    opacity: 0;
    background: ${outlineColor};
    border-radius: 5px;
    cursor: row-resize;
  }

  .row-resize-handle:hover {
    opacity: 1;
  }
`;

const SheetCanvas = styled.canvas`
  position: absolute;
  top: 0px;
  left: 0px;
  width: 100%;
  height: 100%;
`;

const Worksheet = styled.div<{ $hidden: boolean }>`
  display: ${({ $hidden }): string => (!$hidden ? 'block' : 'none')};
  flex-grow: 1;
  position: relative;
  overflow: scroll;
  overscroll-behavior: none;

  ::-webkit-scrollbar-track {
    background-color: rgba(211, 214, 233, 0.2);
    border-radius: 0px;
  }

  ::-webkit-scrollbar-corner {
    background: rgba(211, 214, 233, 0.4);
  }
`;

const Spacer = styled.div`
  position: absolute;
  height: 5000px;
  width: 5000px;
`;
