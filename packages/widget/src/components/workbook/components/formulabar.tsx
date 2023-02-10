import styled from 'styled-components';
import React, { PointerEvent, FunctionComponent, useCallback, useRef } from 'react';
import Editor from '../editor';
import { getSelectedRangeInEditor } from '../editor/util';
import { useWorkbookContext } from '../workbookContext';
import { FocusType, getCellAddress } from '../util';
import { escapeHTML } from '../formulas';

const formulaBarHeight = 30;

type FormulaBarProps = {
  className?: string;
  'data-testid'?: string;
};

export enum FormulaBarTestId {
  CellAddress = 'workbook-formulaBar-cellAddress',
}

const FormulaBar: FunctionComponent<FormulaBarProps> = (properties) => {
  const { model, editorActions, editorState } = useWorkbookContext();
  const { selectedSheet, selectedArea, selectedCell, cellEditing } = editorState;
  const isEditing = cellEditing !== null;
  const formula =
    model?.getFormulaOrValue(selectedSheet, selectedCell.row, selectedCell.column) ?? '';
  const focus = cellEditing?.focus === FocusType.FormulaBar;
  const cursorStart = cellEditing?.cursorStart ?? 0;
  const cursorEnd = cellEditing?.cursorEnd ?? 0;
  const html = cellEditing?.html ?? `<span>${escapeHTML(formula)}</span>`;
  const cellAddress = getCellAddress(selectedArea, selectedCell);

  const formulaBar = useRef<HTMLDivElement>(null);

  const onPointerUp = useCallback(
    (event: PointerEvent) => {
      formulaBar.current?.releasePointerCapture(event.pointerId);
      if (!isEditing) {
        const sel = getSelectedRangeInEditor();
        if (sel) {
          editorActions.onFormulaEditStart(sel);
        }
      }
      event.stopPropagation();
    },
    [isEditing, editorActions],
  );

  const onPointerDown = useCallback(
    (event: PointerEvent) => {
      formulaBar.current?.setPointerCapture(event.pointerId);
    },
    [formulaBar],
  );

  return (
    <FormulaBarContainer data-testid={properties['data-testid']}>
      <NameContainer>
        <CellBarAddress data-testid={FormulaBarTestId.CellAddress}>{cellAddress}</CellBarAddress>
      </NameContainer>
      <FormulaContainer onPointerUp={onPointerUp} onPointerDown={onPointerDown} ref={formulaBar}>
        <Editor
          display
          html={html}
          mode="edit"
          focus={focus}
          cursorStart={cursorStart}
          cursorEnd={cursorEnd}
        />
      </FormulaContainer>
    </FormulaBarContainer>
  );
};

const CellBarAddress = styled.div`
  width: 100%;
  text-align: center;
`;

const FormulaBarContainer = styled.div`
  padding-left: 2px;
  flex-shrink: 0;
  display: flex;
  flex-direction: row;
  align-items: center;
  height: ${formulaBarHeight}px;
`;

const NameContainer = styled.div`
  background: #587af0;
  padding: 4px;
  margin-left: 2px;
  border-radius: 2px;
  color: #ffffff;
  font-style: normal;
  font-weight: normal;
  font-size: 12px;
  display: flex;
  flex-grow: row;
`;
const FormulaContainer = styled.div`
  margin-left: 10px;
  line-height: 22px;
  font-weight: normal;
  width: 100%;
  height: 22px;
  display: flex;
`;

export default FormulaBar;
