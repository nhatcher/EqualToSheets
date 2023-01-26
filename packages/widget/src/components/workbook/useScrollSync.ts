import { RefObject, useEffect, useRef, useCallback } from 'react';
import { ScrollPosition } from './util';

const useScrollSync = (
  worksheetElement: RefObject<HTMLDivElement>,
  scrollPosition: ScrollPosition,
  setScrollPosition: (position: ScrollPosition) => void,
): { onScroll: () => void } => {
  const scrollPositionSyncing = useRef(false);
  // Sync real scrollPosition of worksheetElement with scrollPosition kept in state
  useEffect(() => {
    if (!worksheetElement.current) {
      return;
    }

    if (!scrollPositionSyncing.current) {
      // eslint-disable-next-line no-param-reassign
      worksheetElement.current.scrollLeft = scrollPosition.left;
      // eslint-disable-next-line no-param-reassign
      worksheetElement.current.scrollTop = scrollPosition.top;
    }
    scrollPositionSyncing.current = false;
  }, [scrollPosition, worksheetElement]);

  // Sync scrollPosition kept in state with real scrollPosition of worksheetElement
  const onScroll = useCallback((): void => {
    if (!worksheetElement.current) {
      return;
    }
    const left = worksheetElement.current.scrollLeft;
    const top = worksheetElement.current.scrollTop;

    setScrollPosition({ left, top });
    scrollPositionSyncing.current = true;
  }, [setScrollPosition, worksheetElement]);

  return { onScroll };
};

export default useScrollSync;
