import { KeyboardEvent, useCallback } from 'react';
import { isInReferenceMode, popReferenceFromFormula } from '../formulas';
import { columnNameFromNumber, quoteSheetName } from '../util';
import { useWorkbookContext } from '../workbookContext';

interface Options {
  onMoveCaretToStart: () => void;
  onMoveCaretToEnd: () => void;
  onEditEnd: (delta: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onReferenceCycle: () => void;
  text: string;
  setText: (text: string) => void;
}

const useEditorKeydown = (options: Options): ((event: KeyboardEvent) => void) => {
  const { model, lastRow, lastColumn, editorState } = useWorkbookContext();

  const mode = editorState.cellEditing?.mode ?? 'init';

  const handler = useCallback(
    (event: KeyboardEvent) => {
      const { key } = event;

      // Keyboard-based browse mode when editing in "init" mode.
      // To start, navigate to cell using keyboard, type = and click any arrow key.
      // Reference should appear and be moveable.
      if (
        model &&
        editorState.cellEditing !== null &&
        editorState.cellEditing.mode === 'init' &&
        (key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight')
      ) {
        let baseText = options.text;
        let baseReference: [number, number, number] | null = null;

        const tabs = model.getTabs().map((tab) => tab.name);

        const result = popReferenceFromFormula(
          options.text,
          editorState.cellEditing.sheet,
          tabs,
          model.getTokens,
        );

        if (result) {
          baseText = result.remainderText;
          baseReference = [
            result.reference.sheet,
            result.reference.rowStart,
            result.reference.columnStart,
          ];
        } else if (isInReferenceMode(options.text, options.text.length)) {
          baseReference = [
            editorState.cellEditing.sheet,
            editorState.cellEditing.row,
            editorState.cellEditing.column,
          ];
        }

        if (baseReference) {
          const [deltaRow, deltaColumn] = {
            ArrowUp: [-1, 0],
            ArrowDown: [1, 0],
            ArrowLeft: [0, -1],
            ArrowRight: [0, 1],
          }[key];

          const [sheet, row, column] = baseReference;
          const newReferenceRow = Math.max(1, Math.min(row + deltaRow, lastRow));
          const newReferenceColumn = Math.max(1, Math.min(column + deltaColumn, lastColumn));

          options.setText(
            baseText +
              (sheet !== editorState.cellEditing.sheet ? `${quoteSheetName(tabs[sheet])}!` : '') +
              columnNameFromNumber(newReferenceColumn) +
              newReferenceRow,
          );

          event.stopPropagation();
          event.preventDefault();
          return;
        }
      }

      switch (key) {
        case 'Enter':
          options.onEditEnd({ deltaRow: 1, deltaColumn: 0 });
          event.preventDefault();
          event.stopPropagation();
          break;
        case 'ArrowUp': {
          if (mode === 'init') {
            options.onEditEnd({ deltaRow: -1, deltaColumn: 0 });
          } else {
            options.onMoveCaretToStart();
          }
          event.preventDefault();
          event.stopPropagation();
          break;
        }
        case 'ArrowDown': {
          if (mode === 'init') {
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
          if (mode === 'init') {
            options.onEditEnd({ deltaRow: 0, deltaColumn: -1 });
            event.preventDefault();
            event.stopPropagation();
          }

          break;
        }
        case 'ArrowRight': {
          if (mode === 'init') {
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
    [mode, model, options, editorState, lastRow, lastColumn],
  );
  return handler;
};

export default useEditorKeydown;
