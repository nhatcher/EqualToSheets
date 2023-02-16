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
import RowContextMenuContent from './rowContextMenuContent';
import { isInReferenceMode } from './formulas';
import { useWorkbookContext } from './workbookContext';

export type WorkbookElements = {
  cellEditor: HTMLDivElement | null;
};

export const CLIPBOARD_ID_SESSION_STORAGE_KEY = 'equalTo_clipboardId';

const getNewClipboardId = () => new Date().toISOString();

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
  const contextMenuAnchorElement = useRef<HTMLDivElement>(null);
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

  const onInsertRow = (row: number): void => {
    model?.insertRow(selectedSheet, row);
    setIsRowContextMenuOpen(false);
  };

  const onDeleteRow = (row: number): void => {
    model?.deleteRow(selectedSheet, row);
    setIsRowContextMenuOpen(false);
  };

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

  const onPaste = (event: React.ClipboardEvent) => {
    if (!model) {
      return;
    }
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
    if (mimeType === 'application/json') {
      // We are copying from within the application
      const targetArea = {
        sheet: selectedSheet,
        ...selectedArea,
      };

      try {
        const source = JSON.parse(value);
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
    } else if (mimeType === 'text/plain') {
      model.pasteText(selectedSheet, selectedCell, value);
    } else {
      // NOT IMPLEMENTED
    }
    event.preventDefault();
    event.stopPropagation();
  };

  const onCut = (event: React.ClipboardEvent<Element>) => {
    if (!model) {
      return;
    }
    const { tsv, area, sheetData } = model.copy({ sheet: selectedSheet, ...selectedArea });
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
    event.preventDefault();
    event.stopPropagation();
    // FIXME: It doesn't really work, it should deleteCells too
  };

  // Other
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
  // FIXME: Copy only works for areas [<top,left>, <bottom,right>]
  const onCopy = (event: React.ClipboardEvent<Element>) => {
    if (!model) {
      return;
    }
    const { tsv, area, sheetData } = model.copy({ sheet: selectedSheet, ...selectedArea });
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
    event.preventDefault();
    event.stopPropagation();
  };

  const onPointerDownAtCell = (cell: Cell, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    if (cellEditing) {
      const value = cellInput.current?.value ?? '';
      // FIXME: use ref and figure out selection
      if (isInReferenceMode(value, cellEditing.text.length)) {
        editorActions.onEditPointerDown(cell);
        return;
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
      if (!model || model.isRowReadOnly(selectedSheet, row)) {
        return;
      }
      setIsRowContextMenuOpen(true);
      setRowContextMenu(row);
    },
    [model, selectedSheet],
  );

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
      contextMenuAnchorElement,
      onRowContextMenu,
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

  // TODO: Move WorkbookContainer to workbookContext.tsx
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
        onCopy={onCopy}
        onPaste={onPaste}
        onCut={onCut}
        onContextMenu={onContextMenu}
        ref={worksheetElement}
      >
        <SheetCanvas ref={canvasElement} />
        <CellOutline
          ref={cellOutline} // FIXME: Probably should be outside so we don't need to define event handlers
          onPointerDown={(event) => event.stopPropagation()}
          onPointerUp={(event) => event.stopPropagation()}
          onPointerMove={(event) => event.stopPropagation()}
          onCopy={(event) => event.stopPropagation()}
          onCut={(event) => event.stopPropagation()}
          onPaste={(event) => event.stopPropagation()}
        >
          <Editor display={!!cellEditing} />
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
          <ContextMenuAnchorElement ref={contextMenuAnchorElement} />
        </Menu.Trigger>
        <RowContextMenuContent
          isMenuOpen={isRowContextMenuOpen}
          row={rowContextMenu}
          anchorEl={contextMenuAnchorElement.current}
          onDeleteRow={onDeleteRow}
          onInsertRow={onInsertRow}
          onClose={(): void => {
            setIsRowContextMenuOpen(false);
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
  border-radius: 3px;
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
