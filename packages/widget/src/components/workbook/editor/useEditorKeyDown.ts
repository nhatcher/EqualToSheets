import { KeyboardEvent, useCallback, useRef } from 'react';
import { isInReferenceMode, popReferenceFromFormula } from '../formulas';
import { columnNameFromNumber, quoteSheetName } from '../util';
import { useWorkbookContext } from '../workbookContext';
import { workbookLastColumn, workbookLastRow } from '../constants';

interface Options {
  onMoveCaretToStart: () => void;
  onMoveCaretToEnd: () => void;
  onEditEnd: (delta: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onReferenceCycle: () => void;
  text: string;
  setText: (text: string) => void;
}

const useEditorKeydown = (
  options: Options,
): {
  onKeyDown: (event: KeyboardEvent) => void;
  onCellEditingChange: () => void;
} => {
  const { model, lastRow, lastColumn, editorState } = useWorkbookContext();

  const mode = editorState.cellEditing?.mode ?? 'init';

  const rangeReferenceCellRef = useRef<{ sheet: number; row: number; column: number } | null>(null);

  const onKeyDown = useCallback(
    (event: KeyboardEvent) => {
      const { key, shiftKey } = event;

      /**
       * Keyboard-based browse mode when editing in "init" mode.
       * To start, navigate to cell using keyboard, type = and click any arrow key.
       * Reference should appear and be moveable.
       * @returns `true` if event handled
       */
      const handleKeyboardCellEditing = (): boolean => {
        if (
          model &&
          editorState.cellEditing !== null &&
          editorState.cellEditing.mode === 'init' &&
          (key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight')
        ) {
          let baseText = options.text;
          let baseReference: {
            sheet: number;
            rowStart: number;
            rowEnd: number;
            columnStart: number;
            columnEnd: number;
          } | null = null;

          const tabs = model.getTabs().map((tab) => tab.name);

          const result = popReferenceFromFormula(
            options.text,
            editorState.cellEditing.sheet,
            tabs,
            model.getTokens,
          );

          if (result) {
            baseText = result.remainderText;
            baseReference = {
              sheet: result.reference.sheet,
              rowStart: result.reference.rowStart,
              rowEnd: result.reference.rowEnd,
              columnStart: result.reference.columnStart,
              columnEnd: result.reference.columnEnd,
            };
          } else if (isInReferenceMode(options.text, options.text.length)) {
            baseReference = {
              sheet: editorState.cellEditing.sheet,
              rowStart: editorState.cellEditing.row,
              rowEnd: editorState.cellEditing.row,
              columnStart: editorState.cellEditing.column,
              columnEnd: editorState.cellEditing.column,
            };

            rangeReferenceCellRef.current = {
              sheet: editorState.cellEditing.sheet,
              row: editorState.cellEditing.row,
              column: editorState.cellEditing.column,
            };
          }

          if (baseReference === null) {
            return false;
          }

          const rangeReferenceCell = rangeReferenceCellRef.current;
          if (rangeReferenceCell === null) {
            // Can happen when you type for example '=A1' and hit right arrow key.
            // In that case, fall back to normal navigation.
            return false;
          }

          let outputReference: typeof baseReference;
          if (shiftKey) {
            if (key === 'ArrowUp') {
              if (rangeReferenceCell.row < baseReference.rowEnd) {
                outputReference = {
                  ...baseReference,
                  rowEnd: Math.max(1, baseReference.rowEnd - 1),
                };
              } else if (baseReference.rowStart <= rangeReferenceCell.row) {
                outputReference = {
                  ...baseReference,
                  rowStart: Math.max(1, baseReference.rowStart - 1),
                };
              } else {
                throw new Error(
                  'Could not resize selection. Anchor is not part of the reference range.',
                );
              }
            } else if (key === 'ArrowDown') {
              if (baseReference.rowStart < rangeReferenceCell.row) {
                outputReference = {
                  ...baseReference,
                  rowStart: Math.min(workbookLastRow, baseReference.rowStart + 1),
                };
              } else if (rangeReferenceCell.row <= baseReference.rowEnd) {
                outputReference = {
                  ...baseReference,
                  rowEnd: Math.min(workbookLastRow, baseReference.rowEnd + 1),
                };
              } else {
                throw new Error(
                  'Could not resize selection. Anchor is not part of the reference range.',
                );
              }
            } else if (key === 'ArrowLeft') {
              if (rangeReferenceCell.column < baseReference.columnEnd) {
                outputReference = {
                  ...baseReference,
                  columnEnd: Math.max(0, baseReference.columnEnd - 1),
                };
              } else if (baseReference.columnStart <= rangeReferenceCell.column) {
                outputReference = {
                  ...baseReference,
                  columnStart: Math.max(0, baseReference.columnStart - 1),
                };
              } else {
                throw new Error(
                  'Could not resize selection. Anchor is not part of the reference range.',
                );
              }
            } else if (key === 'ArrowRight') {
              if (baseReference.columnStart < rangeReferenceCell.column) {
                outputReference = {
                  ...baseReference,
                  columnStart: Math.min(workbookLastColumn, baseReference.columnStart + 1),
                };
              } else if (rangeReferenceCell.column <= baseReference.columnEnd) {
                outputReference = {
                  ...baseReference,
                  columnEnd: Math.min(workbookLastColumn, baseReference.columnEnd + 1),
                };
              } else {
                throw new Error(
                  'Could not resize selection. Anchor is not part of the reference range.',
                );
              }
            } else {
              const unknownKey: never = key;
              throw new Error(`Unknown direction key: "${unknownKey}".`);
            }
          } else {
            const [deltaRow, deltaColumn] = {
              ArrowUp: [-1, 0],
              ArrowDown: [1, 0],
              ArrowLeft: [0, -1],
              ArrowRight: [0, 1],
            }[key];

            const { sheet, rowStart: row, columnStart: column } = baseReference;
            const newReferenceRow = Math.max(1, Math.min(row + deltaRow, lastRow));
            const newReferenceColumn = Math.max(1, Math.min(column + deltaColumn, lastColumn));

            outputReference = {
              sheet,
              rowStart: newReferenceRow,
              rowEnd: newReferenceRow,
              columnStart: newReferenceColumn,
              columnEnd: newReferenceColumn,
            };

            rangeReferenceCellRef.current = {
              sheet,
              row: newReferenceRow,
              column: newReferenceColumn,
            };
          }

          const sheetPrefix =
            outputReference.sheet !== editorState.cellEditing.sheet
              ? `${quoteSheetName(tabs[outputReference.sheet])}!`
              : '';

          if (
            outputReference.rowStart === outputReference.rowEnd &&
            outputReference.columnStart === outputReference.columnEnd
          ) {
            options.setText(
              baseText +
                sheetPrefix +
                columnNameFromNumber(outputReference.columnStart) +
                outputReference.rowStart,
            );
          } else {
            options.setText(
              `${baseText}${sheetPrefix}${columnNameFromNumber(outputReference.columnStart)}${
                outputReference.rowStart
              }:${columnNameFromNumber(outputReference.columnEnd)}${outputReference.rowEnd}`,
            );
          }

          event.stopPropagation();
          event.preventDefault();
          return true;
        }

        return false;
      };

      if (handleKeyboardCellEditing()) {
        return;
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

  const onCellEditingChange = useCallback(() => {
    rangeReferenceCellRef.current = null;
  }, []);

  return { onKeyDown, onCellEditingChange };
};

export default useEditorKeydown;
