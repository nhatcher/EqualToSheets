import { useCallback, RefObject } from 'react';
import WorksheetCanvas from './canvas';

type UseResize = (
  worksheetElement: RefObject<HTMLElement>,
  worksheetCanvas: RefObject<WorksheetCanvas>,
) => {
  getCanvasSize: () => { width: number; height: number } | null;
  onResize: () => void;
};

const useResize: UseResize = (worksheetElement, worksheetCanvas) => {
  const getCanvasSize = useCallback((): { width: number; height: number } | null => {
    if (!worksheetElement.current) {
      return null;
    }
    const scrollbarWidth =
      worksheetElement.current.offsetWidth - worksheetElement.current.clientWidth;
    const scrollbarHeight =
      worksheetElement.current.offsetHeight - worksheetElement.current.clientHeight;

    const rect = worksheetElement.current.getBoundingClientRect();

    return {
      width: rect.width - scrollbarWidth,
      height: rect.height - scrollbarHeight,
    };
  }, [worksheetElement]);

  const onResize = useCallback((): void => {
    if (!worksheetCanvas.current) {
      return;
    }

    const canvasSize = getCanvasSize();

    if (!canvasSize) {
      return;
    }

    worksheetCanvas.current.setSize(canvasSize);

    worksheetCanvas.current.renderSheet();
  }, [getCanvasSize, worksheetCanvas]);

  return { getCanvasSize, onResize };
};

export default useResize;
