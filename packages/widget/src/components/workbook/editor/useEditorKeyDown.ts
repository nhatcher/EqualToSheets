import { useCallback, KeyboardEvent } from 'react';
import { CellEditMode } from '../util';

interface Options {
  onMoveCaretToStart: () => void;
  onMoveCaretToEnd: () => void;
  onEditEnd: (delta: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onReferenceCycle: () => void;
  mode: CellEditMode;
}

const useEditorKeydown = (options: Options): ((event: KeyboardEvent) => void) => {
  const handler = useCallback(
    (event: KeyboardEvent) => {
      const { key } = event;
      switch (key) {
        case 'Enter':
          options.onEditEnd({ deltaRow: 1, deltaColumn: 0 });
          event.preventDefault();
          event.stopPropagation();
          break;
        case 'ArrowUp': {
          if (options.mode === 'init') {
            options.onEditEnd({ deltaRow: -1, deltaColumn: 0 });
          } else {
            options.onMoveCaretToStart();
          }
          event.preventDefault();
          event.stopPropagation();
          break;
        }
        case 'ArrowDown': {
          if (options.mode === 'init') {
            options.onEditEnd({ deltaRow: 1, deltaColumn: 0 });
          } else {
            options.onMoveCaretToEnd();
          }
          event.preventDefault();
          event.stopPropagation();
          break;
        }
        case 'Tab': {
          if (event.shiftKey) {
            options.onEditEnd({ deltaRow: 0, deltaColumn: -1 });
          } else {
            options.onEditEnd({ deltaRow: 0, deltaColumn: 1 });
          }
          event.preventDefault();
          event.stopPropagation();

          break;
        }
        case 'Escape': {
          options.onEditEscape();
          event.preventDefault();
          event.stopPropagation();

          break;
        }
        case 'ArrowLeft': {
          if (options.mode === 'init') {
            options.onEditEnd({ deltaRow: 0, deltaColumn: -1 });
            event.preventDefault();
            event.stopPropagation();
          }

          break;
        }
        case 'ArrowRight': {
          if (options.mode === 'init') {
            options.onEditEnd({ deltaRow: 0, deltaColumn: 1 });
            event.preventDefault();
            event.stopPropagation();
          }

          break;
        }
        case 'F4': {
          options.onReferenceCycle();
          event.preventDefault();
          event.stopPropagation();

          break;
        }
        default:
          break;
      }
      event.stopPropagation();
    },
    [options],
  );
  return handler;
};

export default useEditorKeydown;
