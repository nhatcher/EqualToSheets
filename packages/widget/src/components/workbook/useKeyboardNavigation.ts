import { useCallback, KeyboardEvent, RefObject } from 'react';
import Model from './model';
import { WorkbookActions } from './useWorkbookActions';
import { WorkbookState } from './useWorkbookReducer';
import { isEditingKey, isNavigationKey } from './util';

interface Options {
  model: Model | null;
  editorState: WorkbookState;
  editorActions: WorkbookActions;
  root: RefObject<HTMLDivElement>;
}

const useKeyboardNavigation = (options: Options): { onKeyDown: (event: KeyboardEvent) => void } => {
  const { model, editorState, editorActions } = options;
  const { selectedSheet, selectedArea } = editorState;

  // TODO: We can probably drop the useCallback
  const onKeyDown = useCallback(
    (event: KeyboardEvent) => {
      const { key } = event;
      const { root } = options;
      // Silence the linter
      if (!root.current) {
        return;
      }
      if (event.target !== root.current) {
        return;
      }
      if (event.metaKey || event.ctrlKey) {
        switch (key) {
          case 'z': {
            model?.undo();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'y': {
            model?.redo();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'b': {
            model?.toggleFontStyle(selectedSheet, selectedArea, 'bold');
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'i': {
            model?.toggleFontStyle(selectedSheet, selectedArea, 'italics');
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'u': {
            model?.toggleFontStyle(selectedSheet, selectedArea, 'underline');
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'a': {
            // TODO: Area selection. CTRL+A should select "continous" area around the selection,
            // if it does exist then whole sheet is selected.
            event.stopPropagation();
            event.preventDefault();
            break;
          }
          default:
            break;
        }
        if (isNavigationKey(key)) {
          editorActions.onNavigationToEdge(key);
          event.stopPropagation();
          event.preventDefault();
        }
        return;
      }
      if (key === 'F2') {
        editorActions.onCellEditStart();
        event.stopPropagation();
        event.preventDefault();
        return;
      }
      if (isEditingKey(key) || key === 'Backspace') {
        const initText = key === 'Backspace' ? '' : key;
        editorActions.onEditKeyPressStart(initText);
        event.stopPropagation();
        event.preventDefault();
        return;
      }
      // Worksheet Navigation
      if (event.shiftKey) {
        if (
          key === 'ArrowRight' ||
          key === 'ArrowLeft' ||
          key === 'ArrowUp' ||
          key === 'ArrowDown'
        ) {
          editorActions.onExpandAreaSelectedKeyboard(key);

          // Shift + Arrows can be used to select content on a page - prevent this.
          event.preventDefault();
          event.stopPropagation();
        } else if (key === 'Tab') {
          editorActions.onArrowLeft();
          event.stopPropagation();
          event.preventDefault();
        }
        return;
      }
      switch (key) {
        case 'ArrowRight':
        case 'Tab': {
          editorActions.onArrowRight();

          break;
        }
        case 'ArrowLeft': {
          editorActions.onArrowLeft();

          break;
        }
        case 'ArrowDown':
        case 'Enter': {
          editorActions.onArrowDown();

          break;
        }
        case 'ArrowUp': {
          editorActions.onArrowUp();

          break;
        }
        case 'End': {
          editorActions.onKeyEnd();

          break;
        }
        case 'Home': {
          editorActions.onKeyHome();

          break;
        }
        case 'Delete': {
          model?.deleteCells(selectedSheet, selectedArea);

          break;
        }
        case 'PageDown': {
          editorActions.onPageDown();

          break;
        }
        case 'PageUp': {
          editorActions.onPageUp();

          break;
        }
        default:
          break;
      }
      event.stopPropagation();
      event.preventDefault();
    },
    [editorActions, model, options, selectedArea, selectedSheet],
  );
  return { onKeyDown };
};

export default useKeyboardNavigation;
