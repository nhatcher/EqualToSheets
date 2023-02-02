import React, {
  FunctionComponent,
  useEffect,
  useRef,
  useMemo,
  useState,
  useCallback,
  useReducer,
} from 'react';
import Loading from 'src/components/uiKit/loading';
import styled from 'styled-components';
import { fonts } from 'src/theme';
import * as Menu from 'src/components/uiKit/menu';
import Navigation from './components/navigation';
import FormulaBar from './components/formulabar';
import Toolbar from './components/toolbar';
import WorksheetCanvas from './canvas';
import { Cell, FocusType, getCellAddress } from './util';
import useKeyboardNavigation from './useKeyboardNavigation';
import Model, { CellStyle } from './model';
import { useCalcModule } from './useCalcModule';
import usePointer from './usePointer';
import useResize from './useResize';
import Editor from './editor';
import { outlineBackgroundColor, outlineColor } from './constants';
import useWorkbookReducer, { defaultSheetState } from './useWorkbookReducer';
import useScrollSync from './useScrollSync';
import useWorkbookActions from './useWorkbookActions';
import RowContextMenuContent from './rowContextMenuContent';
import { escapeHTML, isInReferenceMode } from './formulas';
import { getSelectedRangeInEditor } from './editor/util';

export enum WorkbookTestId {
  WorkbookContainer = 'workbook-container',
  WorkbookToolbar = 'workbook-toolbar',
  WorkbookFormulaBar = 'workbook-formulaBar',
  WorkbookWorksheet = 'workbook-worksheet',
  WorkbookNavigation = 'workbook-navigation',
  WorkbookCanvasWrapper = 'workbook-canvas-wrapper',
  WorkbookCanvas = 'workbook-canvas',
  WorkbookCellEditor = 'workbook-cell-editor',
}

export type WorkbookElements = {
  cellEditor: HTMLDivElement | null;
};

type WorkbookProps = {
  className?: string;
};

export const CLIPBOARD_ID_SESSION_STORAGE_KEY = 'equalTo_clipboardId';

const getNewClipboardId = () => new Date().toISOString();

const Workbook: FunctionComponent<WorkbookProps> = (properties) => {
  const { className } = properties;

  const [requestRenderId, requestRender] = useReducer((x: number) => x + 1, 0);
  const { calcModule } = useCalcModule();

  // TODO: Move to useModel();
  const [model, setModel] = useState<Model | null>(null);
  useEffect(() => {
    if (!calcModule) {
      return;
    }
    if (!model) {
      const newModel = calcModule.newEmpty();
      setModel(newModel);
    }
  }, [model, calcModule]);

  const [
    { selectedSheet, scrollPosition, selectedCell, selectedArea, extendToArea, cellEditing },
    dispatch,
  ] = useWorkbookReducer(model);

  const tabs = model?.getTabs() ?? [];
  const formula =
    model?.getFormulaOrValue(selectedSheet, selectedCell.row, selectedCell.column) ?? '';

  // References to DOM elements
  const canvasElement = useRef<HTMLCanvasElement>(null);
  const worksheetElement = useRef<HTMLDivElement>(null);
  const scrollElement = useRef<HTMLDivElement>(null);
  const rootElement = useRef<HTMLDivElement>(null);
  const spacerElement = useRef<HTMLDivElement>(null);
  const cellOutline = useRef<HTMLDivElement>(null);
  const areaOutline = useRef<HTMLDivElement>(null);
  const cellOutlineHandle = useRef<HTMLDivElement>(null);
  const extendToOutline = useRef<HTMLDivElement>(null);
  const columnResizeGuide = useRef<HTMLDivElement>(null);
  const rowResizeGuide = useRef<HTMLDivElement>(null);
  const contextMenuAnchorElement = useRef<HTMLDivElement>(null);
  const columnHeaders = useRef<HTMLDivElement>(null);

  const worksheetCanvas = useRef<WorksheetCanvas | null>(null);

  const resizeSpacer = useCallback((options: { deltaWidth: number; deltaHeight: number }): void => {
    if (spacerElement.current) {
      const spacerRect = spacerElement.current.getBoundingClientRect();

      const newSpacerWidth = spacerRect.width + options.deltaWidth;
      const newSpacerHeight = spacerRect.height + options.deltaHeight;

      spacerElement.current.style.width = `${newSpacerWidth}px`;
      spacerElement.current.style.height = `${newSpacerHeight}px`;
    }
  }, []);

  const workbookActions = useWorkbookActions(dispatch, {
    worksheetCanvas,
    worksheetElement,
    rootElement,
  });

  const {
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
    onExtendToEnd: updateSelectionExtendToEnd,
    onNavigationToEdge,
    onEditKeyPressStart,
    onExpandAreaSelectedKeyboard,
    onPointerMoveToCell,
    onEditChange,
    onReferenceCycle,
    onEditEscape,
    onCellEditStart,
    onFormulaEditStart,
  } = workbookActions;

  const focusWorkbook = useCallback(() => {
    rootElement.current?.focus();
  }, []);

  const onChange = useCallback(() => {
    requestRender();
    rootElement.current?.focus();
  }, []);
  // TODO: All of these actions will need to be synced to server and call requestRender
  // Can we use some decorator on model for that?
  const onAddBlankSheet = () => {
    model?.addBlankSheet();
    onChange();
  };

  const onSheetRenamed = (sheet: number, newName: string) => {
    model?.renameSheet(sheet, newName);
    onChange();
  };

  const onSheetDeleted = (sheet: number) => {
    model?.deleteSheet(sheet);
    onSheetSelected(0);
    onChange();
  };

  const onSheetColorChanged = (sheet: number, color: string) => {
    model?.setSheetColor(sheet, color);
    onChange();
  };

  const onInsertRow = (row: number): void => {
    model?.insertRow(selectedSheet, row);
    setIsRowContextMenuOpen(false);
    onChange();
  };

  const onDeleteRow = (row: number): void => {
    model?.deleteRow(selectedSheet, row);
    setIsRowContextMenuOpen(false);
    onChange();
  };

  const onCellsDeleted = () => {
    model?.deleteCells(selectedSheet, selectedArea);
    onChange();
  };

  const onUndo = () => {
    model?.undo();
    onChange();
  };

  const onRedo = () => {
    model?.redo();
    onChange();
  };

  const onToggleBold = () => {
    model?.toggleFontStyle(selectedSheet, selectedArea, 'bold');
    onChange();
  };
  const onToggleItalic = () => {
    model?.toggleFontStyle(selectedSheet, selectedArea, 'italic');
    onChange();
  };
  const onToggleUnderline = () => {
    model?.toggleFontStyle(selectedSheet, selectedArea, 'underline');
    onChange();
  };
  const onToggleStrike = () => {
    model?.toggleFontStyle(selectedSheet, selectedArea, 'strikethrough');
    onChange();
  };
  const onToggleAlignLeft = () => {
    model?.toggleAlign(selectedSheet, selectedArea, 'left');
    onChange();
  };
  const onToggleAlignRight = () => {
    model?.toggleAlign(selectedSheet, selectedArea, 'right');
    onChange();
  };
  const onToggleAlignCenter = () => {
    model?.toggleAlign(selectedSheet, selectedArea, 'center');
    onChange();
  };
  const onTextColorPicked = (color: string) => {
    model?.setTextColor(selectedSheet, selectedArea, color);
    onChange();
  };
  const onFillColorPicked = (color: string) => {
    model?.setFillColor(selectedSheet, selectedArea, color);
    onChange();
  };
  const onNumberFormatPicked = (numberFormat: string) => {
    model?.setNumberFormat(selectedSheet, selectedArea, numberFormat);
    onChange();
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
      onChange();
    },
    [model, onChange, resizeSpacer],
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
      onChange();
    },
    [model, onChange, resizeSpacer],
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
    onChange();
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
    onChange();
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

  const onExtendToEnd = () => {
    if (!model || !extendToArea) {
      return;
    }
    model.extendTo(selectedSheet, selectedArea, extendToArea);
    updateSelectionExtendToEnd();
    onChange();
  };

  const onSetCellValue = (sheet: number, row: number, column: number, text: string) => {
    if (!model) {
      return;
    }
    model.setCellValue(sheet, row, column, text);
    onChange();
  };

  const onPointerDownAtCell = (cell: Cell, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    if (cellEditing) {
      if (isInReferenceMode(cellEditing.text, cellEditing.cursorEnd)) {
        workbookActions.onEditPointerDown(cell);
        return;
      }
      // FIXME: This is out of context. This happens while you are editing a cell and finish the editing by
      // clicking somewhere else.
      // This probably should be done in the onBlur event in the editor.
      // The cellEditing object might have not been updated because we debounce key strokes,
      // So we use the text in the editor (note that we are editing the cell)
      const sel = getSelectedRangeInEditor();
      const text = sel ? sel.text : '';
      onSetCellValue(cellEditing.sheet, cellEditing.row, cellEditing.column, text);
    }
    workbookActions.onPointerDownAtCell(cell);
    onChange();
  };

  const onEditEnd = (delta: { deltaRow: number; deltaColumn: number }) => {
    if (!cellEditing) {
      return;
    }
    onSetCellValue(cellEditing.sheet, cellEditing.row, cellEditing.column, cellEditing.text);
    workbookActions.onEditEnd(delta);
    onChange();
  };

  const { getCanvasSize, onResize } = useResize(worksheetElement, worksheetCanvas);
  const { onScroll } = useScrollSync(scrollElement, scrollPosition, setScrollPosition);

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
      onAreaSelected,
      onPointerDownAtCell,
      onPointerMoveToCell,
      onExtendToCell,
      onExtendToEnd,
      canvasElement,
      worksheetElement,
      worksheetCanvas,
      contextMenuAnchorElement,
      onRowContextMenu,
    });

  const { onKeyDown: onKeyDownNavigation } = useKeyboardNavigation({
    onArrowDown,
    onArrowUp,
    onArrowLeft,
    onArrowRight,
    onKeyHome,
    onKeyEnd,
    onCellsDeleted,
    onCellEditStart,
    onBold: onToggleBold,
    onItalic: onToggleItalic,
    onUnderline: onToggleUnderline,
    onExpandAreaSelectedKeyboard,
    onUndo,
    onRedo,
    onEditKeyPressStart,
    onNavigationToEdge,
    onPageDown,
    onPageUp,
    root: rootElement,
  });

  const canEdit = useMemo(
    () => !(model?.isCellReadOnly(selectedSheet, selectedCell.row, selectedCell.column) ?? true),
    [model, selectedCell.column, selectedCell.row, selectedSheet],
  );

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
  }, [requestRender, getCanvasSize, model, onColumnWidthChange, onResize, onRowHeightChange]);

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
  }, [extendToArea, requestRender, scrollPosition, selectedArea, selectedCell, selectedSheet]);

  // Sync canvas with editing state
  useEffect(() => {
    worksheetCanvas.current?.setCellEditing(cellEditing);
    requestRender();
  }, [cellEditing, requestRender]);

  useEffect(() => {
    worksheetCanvas.current?.renderSheet();
  }, [
    requestRenderId, // enables rerendering on request
  ]);

  let cellStyle: CellStyle = {
    horizontal_alignment: 'default',
    read_only: true,
    quote_prefix: false,
    num_fmt: 'general',
    fill: { pattern_type: 'solid', fg_color: { RGB: '#e5e5e5' }, bg_color: { RGB: '#000000' } },
    font: {
      strike: false,
      u: false,
      b: false,
      i: false,
      sz: 11,
      color: { RGB: '#000000' },
      name: 'Calibri',
      family: 2,
      scheme: 'minor',
    },
    border: {
      left: { style: 'none', color: { RGB: '#FFFFF' } },
      right: { style: 'none', color: { RGB: '#FFFFF' } },
      top: { style: 'none', color: { RGB: '#FFFFF' } },
      bottom: { style: 'none', color: { RGB: '#FFFFF' } },
      diagonal: { style: 'none', color: { RGB: '#FFFFF' } },
    },
  };
  if (model) {
    cellStyle = model.getCellStyle(selectedSheet, selectedCell.row, selectedCell.column);
  }

  const cellAddress = getCellAddress(selectedArea, selectedCell);

  return (
    <WorkbookContainer
      className={className}
      ref={rootElement}
      data-testid={WorkbookTestId.WorkbookContainer}
      tabIndex={0}
      onKeyDown={onKeyDownNavigation}
      onContextMenu={(event): void => {
        // prevents the browser menu
        event.preventDefault();
      }}
    >
      <Toolbar
        canUndo={model === null ? false : model.canUndo()}
        canRedo={model === null ? false : model.canRedo()}
        canEdit={canEdit}
        onUndo={onUndo}
        onRedo={onRedo}
        onToggleBold={onToggleBold}
        onToggleItalic={onToggleItalic}
        onToggleUnderline={onToggleUnderline}
        onToggleStrike={onToggleStrike}
        onToggleAlignLeft={onToggleAlignLeft}
        onToggleAlignCenter={onToggleAlignCenter}
        onToggleAlignRight={onToggleAlignRight}
        onTextColorPicked={onTextColorPicked}
        onFillColorPicked={onFillColorPicked}
        onNumberFormatPicked={onNumberFormatPicked}
        focusWorkbook={focusWorkbook}
        fontColor={cellStyle.font.color.RGB}
        fillColor={cellStyle.fill.fg_color?.RGB ?? '#FFFFFF'}
        bold={!!cellStyle.font.b}
        italic={!!cellStyle.font.i}
        underline={!!cellStyle.font.u}
        strike={!!cellStyle.font.strike}
        alignment={cellStyle.horizontal_alignment}
        numFmt={cellStyle.num_fmt}
        data-testid={WorkbookTestId.WorkbookToolbar}
        selectedArea={selectedArea}
      />
      <FormulaBar
        cellAddress={cellAddress}
        data-testid={WorkbookTestId.WorkbookFormulaBar}
        onEditStart={onFormulaEditStart}
        onEditChange={onEditChange}
        onEditEnd={onEditEnd}
        onEditEscape={onEditEscape}
        onReferenceCycle={onReferenceCycle}
        html={cellEditing?.html || `<span>${escapeHTML(formula)}</span>`}
        cursorStart={cellEditing?.cursorStart || 0}
        cursorEnd={cellEditing?.cursorEnd || 0}
        focus={cellEditing?.focus === FocusType.FormulaBar}
        isEditing={cellEditing !== null}
      />
      <Menu.Root open={isRowContextMenuOpen} onOpenChange={setIsRowContextMenuOpen}>
        <Menu.Trigger>
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
      <Worksheet
        data-testid={WorkbookTestId.WorkbookWorksheet}
        $hidden={!model}
        onScroll={onScroll}
        ref={scrollElement}
      >
        <Spacer ref={spacerElement} />
        <SheetCanvasWrapper
          data-testid={WorkbookTestId.WorkbookCanvasWrapper}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onDoubleClick={(event) => {
            onCellEditStart();
            event.stopPropagation();
            event.preventDefault();
          }}
          onCopy={onCopy}
          onPaste={onPaste}
          onCut={onCut}
          onContextMenu={onContextMenu}
          ref={worksheetElement}
        >
          <SheetCanvas ref={canvasElement} data-testid={WorkbookTestId.WorkbookCanvas} />
          <CellOutline ref={cellOutline}>
            <Editor
              data-testid={WorkbookTestId.WorkbookCellEditor}
              onEditChange={onEditChange}
              onEditEnd={onEditEnd}
              onEditEscape={onEditEscape}
              onReferenceCycle={onReferenceCycle}
              display={!!cellEditing}
              focus={cellEditing?.focus === FocusType.Cell}
              html={cellEditing?.html ?? ''}
              cursorStart={cellEditing?.cursorStart ?? 0}
              cursorEnd={cellEditing?.cursorEnd ?? 0}
              mode={cellEditing?.mode ?? 'init'}
            />
          </CellOutline>
          <AreaOutline ref={areaOutline} />
          <ExtendToOutline ref={extendToOutline} />
          <CellOutlineHandle ref={cellOutlineHandle} onPointerDown={onPointerHandleDown} />
          <ColumnResizeGuide ref={columnResizeGuide} />
          <RowResizeGuide ref={rowResizeGuide} />
          <ColumnHeaders ref={columnHeaders} />
        </SheetCanvasWrapper>
      </Worksheet>
      {!model && <Loading marginTop="30px" />}
      <Navigation
        data-testid={WorkbookTestId.WorkbookNavigation}
        tabs={tabs}
        selectedSheet={selectedSheet}
        onSheetSelected={onSheetSelected}
        onAddBlankSheet={onAddBlankSheet}
        onSheetColorChanged={onSheetColorChanged}
        disabled={!model}
        onSheetRenamed={onSheetRenamed}
        onSheetDeleted={onSheetDeleted}
      />
    </WorkbookContainer>
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

// TODO: Unset css that could affect workbook
// TODO: Have to set background color, because there are some empty spots, sic!
const WorkbookContainer = styled.div`
  text-align: left;
  background-color: #fff;
  display: flex;
  flex-direction: column;
  height: 100%;
  width: 100%;
  font-family: ${fonts.mono};
  color: #000;
  font-size: 16px;
  &:focus {
    outline: none;
  }
`;
