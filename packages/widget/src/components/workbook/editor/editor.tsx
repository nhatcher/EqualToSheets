import styled, { css } from 'styled-components';
import React, { FunctionComponent, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { palette } from 'src/theme';
import useEditorKeyDown from './useEditorKeyDown';
import { useWorkbookContext } from '../workbookContext';
import { FocusType } from '../util';

export enum EditorPageTestId {
  FormulaEditor = 'workbook-editor-formula-editor',
}

function formulaToHTML(formula: string) {
  return formula
    .split('')
    .map((char: string, index: number) =>
      index % 2 === 0
        ? `<span style="color:red;">${char}</span>`
        : `<span style="color:blue;">${char}</span>`,
    )
    .join('');
}

const Editor: FunctionComponent<{
  display: boolean;
}> = (properties) => {
  const { model, editorActions, editorState, formulaBarEditor, cellInput } = useWorkbookContext();
  const { selectedSheet, selectedCell, cellEditing } = editorState;
  const { onEditEscape } = editorActions;
  const [text, setText] = useState('');
  const html = formulaToHTML(text);
  const barInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const formula =
      model?.getFormulaOrValue(selectedSheet, selectedCell.row, selectedCell.column) ?? '';
    setText(formula);
  }, [model, selectedCell.column, selectedCell.row, selectedSheet]);

  // TODO: onReferenceCycle not needed for demo
  const onReferenceCycle = () => {};
  const onEditEnd = (delta: Parameters<(typeof editorActions)['onEditEnd']>[1]) => {
    editorActions.onEditEnd(text, delta);
  };

  useEffect(() => {
    if (cellEditing) {
      setText(cellEditing.text);
      if (cellEditing.focus === FocusType.Cell) {
        cellInput.current?.focus();
      } else {
        barInputRef.current?.focus();
      }
    }
  }, [cellEditing, cellInput]);

  const cellEditorKeyDown = useEditorKeyDown({
    onEditEnd,
    onEditEscape,
    onReferenceCycle,
    mode: cellEditing?.text === cellInput.current?.value ? 'init' : 'edit',
  });

  return (
    <>
      <CellEditorContainer $display={properties.display}>
        <MaskContainer dangerouslySetInnerHTML={{ __html: html }} />
        <input
          ref={cellInput}
          spellCheck="false"
          value={text}
          onChange={(event) => setText(event.target.value)}
          onKeyDown={cellEditorKeyDown}
        />
      </CellEditorContainer>
      {formulaBarEditor.current
        ? createPortal(
            <CellEditorContainer $display>
              <MaskContainer dangerouslySetInnerHTML={{ __html: html }} />
              <input
                ref={barInputRef}
                onFocus={() => {
                  if (!cellEditing) {
                    editorActions.onFormulaEditStart();
                  }
                }}
                spellCheck="false"
                value={text}
                onChange={(event) => setText(event.target.value)}
                onKeyDown={cellEditorKeyDown}
              />
            </CellEditorContainer>,
            formulaBarEditor.current,
          )
        : null}
    </>
  );
};

export default Editor;

const EditorFontCSS = css`
  line-height: 22px;
  font-weight: normal;
  height: 22px;
  flex-grow: 1;
  font-family: 'Fira Mono', 'Adjusted Courier New Fallback', serif;
  font-size: 16px;
  padding: 0px;
`;
const CellEditorContainer = styled.div<{ $display: boolean }>`
  display: ${({ $display }): string => ($display ? 'flex' : 'none')};
  box-sizing: border-box;
  position: relative;
  width: 100%;
  padding: 0px;
  input {
    color: transparent;
    caret-color: ${palette.text.primary};
    outline: none;
    border: none;
    ${EditorFontCSS}
  }
  border-width: 0px;
  outline: none;
  resize: none;
  white-space: pre-wrap;
  vertical-align: bottom;
  overflow: hidden;
`;

const MaskContainer = styled.div`
  position: absolute;
  pointer-events: none;
  ${EditorFontCSS}
`;
