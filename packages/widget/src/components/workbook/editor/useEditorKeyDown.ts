import { useCallback, KeyboardEvent } from 'react';
import { CellEditMode } from '../util';
import { getSelectedRangeInEditor } from './util';

interface Options {
  onEditEnd: (delta: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onReferenceCycle: (text: string, cursorStart: number, cursorEnd: number) => void;
  mode: CellEditMode;
}

const useEditorKeydown = (options: Options): ((event: KeyboardEvent) => void) => {
  const handler = useCallback(
    (event: KeyboardEvent) => {
      const { key } = event;
      switch (key) {
        case 'Enter':
        case 'ArrowDown': {
          options.onEditEnd({ deltaRow: 1, deltaColumn: 0 });
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
        case 'ArrowUp': {
          options.onEditEnd({ deltaRow: -1, deltaColumn: 0 });
          event.preventDefault();
          event.stopPropagation();

          break;
        }
        case 'F4': {
          const caret = getSelectedRangeInEditor();
          if (!caret) {
            break;
          }
          const cursorEnd = caret.end;
          const cursorStart = caret.start;
          options.onReferenceCycle(caret.text, cursorStart, cursorEnd);
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
