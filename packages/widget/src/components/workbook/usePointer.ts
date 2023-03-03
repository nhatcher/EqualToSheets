import { useCallback, RefObject, PointerEvent, useRef } from 'react';
import WorksheetCanvas, { headerColumnWidth, headerRowHeight } from './canvas';
import { Area, Cell, Border } from './util';

interface PointerSettings {
  canvasElement: RefObject<HTMLCanvasElement>;
  worksheetCanvas: RefObject<WorksheetCanvas | null>;
  worksheetElement: RefObject<HTMLDivElement>;
  rowContextMenuAnchorElement: RefObject<HTMLDivElement>;
  columnContextMenuAnchorElement: RefObject<HTMLDivElement>;
  onPointerDownAtCell: (cell: Cell, event: React.MouseEvent) => void;
  onPointerMoveToCell: (cell: Cell) => void;
  onAreaSelected: (area: Area, border: Border) => void;
  onExtendToCell: (cell: Cell) => void;
  onExtendToEnd: () => void;
  onRowContextMenu: (row: number) => void;
  onColumnContextMenu: (column: number) => void;
}

interface PointerEvents {
  onPointerDown: (event: PointerEvent) => void;
  onPointerMove: (event: PointerEvent) => void;
  onPointerUp: (event: PointerEvent) => void;
  onPointerHandleDown: (event: PointerEvent) => void;
  onContextMenu: (event: React.MouseEvent) => void;
}

const usePointer = (options: PointerSettings): PointerEvents => {
  const isSelecting = useRef(false);
  const isExtending = useRef(false);

  const onContextMenu = useCallback(
    (event: React.MouseEvent): void => {
      let x = event.clientX;
      let y = event.clientY;
      const {
        canvasElement,
        worksheetElement,
        worksheetCanvas,
        onRowContextMenu,
        rowContextMenuAnchorElement,
        onColumnContextMenu,
        columnContextMenuAnchorElement,
      } = options;
      const worksheet = worksheetCanvas.current;
      const canvas = canvasElement.current;
      const worksheetWrapper = worksheetElement.current;
      // Silence the linter
      if (!canvas || !worksheet || !worksheetWrapper) {
        return;
      }
      const canvasRect = canvas.getBoundingClientRect();
      x -= canvasRect.x;
      y -= canvasRect.y;
      const menuAnchorOffsetY = 10;
      if (x > 0 && x < headerColumnWidth && y > headerRowHeight && y < canvasRect.height) {
        // Click on a row number
        const cell = worksheet.getCellByCoordinates(headerColumnWidth, y);
        if (cell) {
          event.preventDefault();
          event.stopPropagation();
          if (rowContextMenuAnchorElement.current) {
            const scrollPosition = worksheet.getScrollPosition();
            rowContextMenuAnchorElement.current.style.left = `${x + scrollPosition.left}px`;
            rowContextMenuAnchorElement.current.style.top = `${
              y + scrollPosition.top + menuAnchorOffsetY
            }px`;
          }
          options.onPointerDownAtCell(cell, event);
          onRowContextMenu(cell.row);
        }
      }
      if (x > headerColumnWidth && x < canvas.width && y > 0 && y < headerRowHeight) {
        // Click on a column number
        const cell = worksheet.getCellByCoordinates(x, headerRowHeight);
        if (cell) {
          event.preventDefault();
          event.stopPropagation();
          if (columnContextMenuAnchorElement.current) {
            const scrollPosition = worksheet.getScrollPosition();
            columnContextMenuAnchorElement.current.style.left = `${x + scrollPosition.left}px`;
            columnContextMenuAnchorElement.current.style.top = `${
              y + scrollPosition.top + menuAnchorOffsetY
            }px`;
          }
          options.onPointerDownAtCell(cell, event);
          onColumnContextMenu(cell.column);
        }
      }
    },
    [options],
  );

  const onPointerMove = useCallback(
    (event: PointerEvent): void => {
      // Range selections are disabled on non-mouse devices. Use touch move only
      // to scroll for now.
      if (event.pointerType !== 'mouse') {
        return;
      }

      if (isSelecting.current) {
        const { canvasElement, worksheetCanvas } = options;
        const canvas = canvasElement.current;
        const worksheet = worksheetCanvas.current;
        // Silence the linter
        if (!worksheet || !canvas) {
          return;
        }
        let x = event.clientX;
        let y = event.clientY;
        const canvasRect = canvas.getBoundingClientRect();
        x -= canvasRect.x;
        y -= canvasRect.y;
        const cell = worksheet.getCellByCoordinates(x, y);
        if (cell) {
          options.onPointerMoveToCell(cell);
        }
      } else if (isExtending.current) {
        const { canvasElement, worksheetCanvas } = options;
        const canvas = canvasElement.current;
        const worksheet = worksheetCanvas.current;
        // Silence the linter
        if (!worksheet || !canvas) {
          return;
        }
        let x = event.clientX;
        let y = event.clientY;
        const canvasRect = canvas.getBoundingClientRect();
        x -= canvasRect.x;
        y -= canvasRect.y;
        const cell = worksheet.getCellByCoordinates(x, y);
        if (!cell) {
          return;
        }
        options.onExtendToCell(cell);
      }
    },
    [options],
  );

  const onPointerUp = useCallback(
    (event: PointerEvent): void => {
      if (isSelecting.current) {
        const { worksheetElement } = options;
        isSelecting.current = false;
        worksheetElement.current?.releasePointerCapture(event.pointerId);
      } else if (isExtending.current) {
        const { worksheetElement } = options;
        isExtending.current = false;
        worksheetElement.current?.releasePointerCapture(event.pointerId);
        options.onExtendToEnd();
      }
    },
    [options],
  );

  const onPointerDown = useCallback(
    (event: PointerEvent) => {
      let x = event.clientX;
      let y = event.clientY;
      const { canvasElement, worksheetElement, worksheetCanvas } = options;
      const worksheet = worksheetCanvas.current;
      const canvas = canvasElement.current;
      const worksheetWrapper = worksheetElement.current;
      // Silence the linter
      if (!canvas || !worksheet || !worksheetWrapper) {
        return;
      }
      const canvasRect = canvas.getBoundingClientRect();
      x -= canvasRect.x;
      y -= canvasRect.y;
      // Makes sure is in the sheet area
      if (
        x > canvasRect.width ||
        x < headerColumnWidth ||
        y < headerRowHeight ||
        y > canvasRect.height
      ) {
        if (x > 0 && x < headerColumnWidth && y > headerRowHeight && y < canvasRect.height) {
          // Click on a row number
          const cell = worksheet.getCellByCoordinates(headerColumnWidth, y);
          if (cell) {
            // TODO
            // Row selected
          }
        }
        return;
      }
      const cell = worksheet.getCellByCoordinates(x, y);
      if (cell) {
        options.onPointerDownAtCell(cell, event);
        isSelecting.current = true;
        worksheetWrapper.setPointerCapture(event.pointerId);
      }
    },
    [options],
  );

  const onPointerHandleDown = useCallback(
    (event: PointerEvent) => {
      const worksheetWrapper = options.worksheetElement.current;
      // Silence the linter
      if (!worksheetWrapper) {
        return;
      }
      isExtending.current = true;
      worksheetWrapper.setPointerCapture(event.pointerId);
      event.stopPropagation();
      event.preventDefault();
    },
    [options],
  );

  return {
    onPointerDown,
    onPointerMove,
    onPointerUp,
    onPointerHandleDown,
    onContextMenu,
  };
};

export default usePointer;
