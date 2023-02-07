import styled from 'styled-components';
import React, { FunctionComponent, useEffect, useMemo, useRef } from 'react';
import { debounce } from 'src/util';
import useEditorKeyDown from './useEditorKeyDown';
import setCaretPosition, { editorClass, getSelectedRangeInEditor } from './util';
import { CellEditMode } from '../util';
import { useWorkbookContext } from '../workbookContext';

export enum EditorPageTestId {
  FormulaEditor = 'workbook-editor-formula-editor',
}
export interface EditorProps {
  display: boolean;
  html: string;
  cursorStart: number;
  cursorEnd: number;
  focus: boolean;
  mode: CellEditMode;
}

const Editor: FunctionComponent<EditorProps> = (properties) => {
  const { editorActions } = useWorkbookContext();
  const { onReferenceCycle, onEditChange, onEditEscape } = editorActions;
  const cellEditorElement = useRef<HTMLDivElement>(null);

  const onEditEnd = (delta: { deltaRow: number; deltaColumn: number }) => {
    const sel = getSelectedRangeInEditor();
    if (sel) {
      onEditChange(sel.text, sel.start, sel.end);
    }
    editorActions.onEditEnd(delta);
  };

  const cellEditorKeyDown = useEditorKeyDown({
    onEditEnd,
    onEditEscape,
    onReferenceCycle,
    mode: properties.mode,
  });

  useEffect(() => {
    const editor = cellEditorElement.current;
    if (editor) {
      if (properties.display) {
        const { html, focus, cursorStart, cursorEnd } = properties;
        editor.innerHTML = html;
        if (focus) {
          if (cursorStart <= cursorEnd) {
            setCaretPosition(editor, { start: cursorStart, end: cursorEnd });
          } else {
            setCaretPosition(editor, { start: cursorEnd, end: cursorStart });
          }
          editor.focus();
        }
      } else {
        // Clearing the content of the cell editor after being used
        editor.innerHTML = '<span></span><span style="flex-grow: 1;"></span>';
      }
    }
  });

  // We do not want to redraw the canvas with each keystroke if they are in quick succession.
  // See issue: #1407
  const onInput = useMemo(
    () =>
      debounce(() => {
        const sel = getSelectedRangeInEditor();
        if (sel) {
          onEditChange(sel.text, sel.start, sel.end);
        }
      }, 500),
    [onEditChange],
  );

  return (
    <CellEditor
      // FIXME: This ID doesn't seem right, it's used for formula editor and for cell editor
      data-testid={EditorPageTestId.FormulaEditor}
      className={editorClass}
      ref={cellEditorElement}
      contentEditable="true"
      spellCheck="false"
      onKeyDown={cellEditorKeyDown}
      onInput={onInput}
      onPointerDown={(event): void => {
        if (properties.focus) {
          event.stopPropagation();
        }
      }}
      onPointerMove={(event): void => {
        if (properties.focus) {
          event.stopPropagation();
        }
      }}
      onPointerUp={(event): void => {
        if (properties.focus) {
          event.stopPropagation();
        }
      }}
      onCut={(event): void => event.stopPropagation()}
      onCopy={(event): void => event.stopPropagation()}
      onPaste={(event): void => {
        // We need to sanitize the input
        const value = event.clipboardData.getData('text/plain');
        if (!value) {
          event.preventDefault();
          event.stopPropagation();
          return;
        }
        const editor = cellEditorElement.current;
        if (editor) {
          const sel = getSelectedRangeInEditor();
          if (sel) {
            const newEnd = sel.start + value.length;
            const text = sel.text.slice(0, sel.start) + value + sel.text.slice(sel.end);
            onEditChange(text, newEnd, newEnd);
          }
        }
        event.preventDefault();
        event.stopPropagation();
      }}
      $display={properties.display}
    />
  );
};

export default Editor;

const CellEditor = styled.div<{ $display: boolean }>`
  box-sizing: border-box;
  position: relative;
  width: 100%;
  padding: 0px;
  border-width: 0px;
  outline: none;
  resize: none;
  white-space: pre-wrap;
  vertical-align: bottom;
  overflow: hidden;
  display: ${({ $display }): string => ($display ? 'block' : 'none')};
  span {
    min-width: 1px;
  }
`;
