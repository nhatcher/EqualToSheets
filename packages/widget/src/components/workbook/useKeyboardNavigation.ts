import { useCallback, KeyboardEvent, RefObject } from 'react';
import { isEditingKey, isNavigationKey, NavigationKey } from './util';

export enum Border {
  Top = 'top',
  Bottom = 'bottom',
  Right = 'right',
  Left = 'left',
}

interface Options {
  onCellsDeleted: () => void;
  onExpandAreaSelectedKeyboard: (key: 'ArrowRight' | 'ArrowLeft' | 'ArrowUp' | 'ArrowDown') => void;
  onEditKeyPressStart: (initText: string) => void;
  onCellEditStart: () => void;
  onBold: () => void;
  onItalic: () => void;
  onUnderline: () => void;
  onNavigationToEdge: (direction: NavigationKey) => void;
  onPageDown: () => void;
  onPageUp: () => void;
  onArrowDown: () => void;
  onArrowUp: () => void;
  onArrowLeft: () => void;
  onArrowRight: () => void;
  onKeyHome: () => void;
  onKeyEnd: () => void;
  onUndo: () => void;
  onRedo: () => void;
  root: RefObject<HTMLDivElement>;
}

const useKeyboardNavigation = (options: Options): { onKeyDown: (event: KeyboardEvent) => void } => {
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
            options.onUndo();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'y': {
            options.onRedo();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'b': {
            options.onBold();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'i': {
            options.onItalic();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          case 'u': {
            options.onUnderline();
            event.stopPropagation();
            event.preventDefault();

            break;
          }
          default:
            break;
        }
        if (isNavigationKey(key)) {
          options.onNavigationToEdge(key);
          event.stopPropagation();
          event.preventDefault();
        }
        return;
      }
      if (key === 'F2') {
        options.onCellEditStart();
        event.stopPropagation();
        event.preventDefault();
        return;
      }
      if (isEditingKey(key) || key === 'Backspace') {
        const initText = key === 'Backspace' ? '' : key;
        options.onEditKeyPressStart(initText);
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
          options.onExpandAreaSelectedKeyboard(key);
        } else if (key === 'Tab') {
          options.onArrowLeft();
          event.stopPropagation();
          event.preventDefault();
        }
        return;
      }
      switch (key) {
        case 'ArrowRight':
        case 'Tab': {
          options.onArrowRight();

          break;
        }
        case 'ArrowLeft': {
          options.onArrowLeft();

          break;
        }
        case 'ArrowDown':
        case 'Enter': {
          options.onArrowDown();

          break;
        }
        case 'ArrowUp': {
          options.onArrowUp();

          break;
        }
        case 'End': {
          options.onKeyEnd();

          break;
        }
        case 'Home': {
          options.onKeyHome();

          break;
        }
        case 'Delete': {
          options.onCellsDeleted();

          break;
        }
        case 'PageDown': {
          options.onPageDown();

          break;
        }
        case 'PageUp': {
          options.onPageUp();

          break;
        }
        default:
          break;
      }
      event.stopPropagation();
      event.preventDefault();
    },
    [options],
  );
  return { onKeyDown };
};

export default useKeyboardNavigation;
