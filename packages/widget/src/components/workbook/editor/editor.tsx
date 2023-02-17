import styled, { css } from 'styled-components';
import React, { Fragment, FunctionComponent, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { palette } from 'src/theme';
import useEditorKeyDown from './useEditorKeyDown';
import { useWorkbookContext } from '../workbookContext';
import { FocusType } from '../util';
import {
  ColoredFormulaReference,
  getColoredReferences,
  getReferencesFromFormula,
} from '../formulas';

export enum EditorPageTestId {
  FormulaEditor = 'workbook-editor-formula-editor',
}

const Editor: FunctionComponent<{
  display: boolean;
  onReferencesChanged: (references: ColoredFormulaReference[]) => void;
}> = (properties) => {
  const { onReferencesChanged } = properties;
  const { model, editorActions, editorState, formulaBarEditor, cellInput, formulaBarInput } =
    useWorkbookContext();
  const { selectedSheet, selectedCell, cellEditing } = editorState;
  const { onEditEscape } = editorActions;

  const [text, setText] = useState('');

  const formulaReferences = getReferencesFromFormula(
    text,
    selectedSheet,
    model?.getTabs().map((tab) => tab.name) ?? [],
    model?.getTokens ?? (() => []),
  );
  const coloredReferences = getColoredReferences(formulaReferences);

  // Sync references with canvas
  const coloredReferencesJson = useRef(JSON.stringify(coloredReferences));
  useEffect(() => {
    const currentJson = JSON.stringify(coloredReferences);
    if (coloredReferencesJson.current !== currentJson) {
      coloredReferencesJson.current = currentJson;
      onReferencesChanged(coloredReferences);
    }
  }, [coloredReferences, onReferencesChanged]);

  const styledFormula = styleFormula(text, coloredReferences);

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
        formulaBarInput.current?.focus();
      }
    }
  }, [cellEditing, cellInput, formulaBarInput]);

  const cellEditorKeyDown = useEditorKeyDown({
    onEditEnd,
    onEditEscape,
    onReferenceCycle,
    mode: cellEditing?.text === cellInput.current?.value ? 'init' : 'edit',
  });

  return (
    <>
      <CellEditorContainer $display={properties.display}>
        <MaskContainer>{styledFormula}</MaskContainer>
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
              <MaskContainer>{styledFormula}</MaskContainer>
              <input
                ref={formulaBarInput}
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

function styleFormula(formula: string, coloredReferences: ColoredFormulaReference[]) {
  const nodes = [];
  let currentIndex = 0;
  coloredReferences.forEach((reference) => {
    const { start, end, color } = reference;
    const key = JSON.stringify(reference);
    if (reference.start !== currentIndex) {
      nodes.push(<Fragment key={`${key}-prefix`}>{formula.slice(currentIndex, start)}</Fragment>);
    }
    nodes.push(
      <span key={key} style={{ color }}>
        {formula.slice(start, end)}
      </span>,
    );
    currentIndex = end;
  });
  if (currentIndex < formula.length) {
    nodes.push(
      <Fragment key="formula-suffix">{formula.slice(currentIndex, formula.length)}</Fragment>,
    );
  }
  return nodes;
}

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
