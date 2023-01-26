import styled from 'styled-components';
import React, { PointerEvent, FunctionComponent, useCallback, useRef } from 'react';
import { ChevronDownIcon } from 'src/components/uiKit/icons';
import { palette } from 'src/theme';
import Editor from '../editor';
import { headerColumnWidth } from '../canvas';
import { getSelectedRangeInEditor, EditorSelection } from '../editor/util';

const formulaBarHeight = 30;

type FormulaBarProps = {
  className?: string;
  'data-testid'?: string;
  cellAddress: string;
  onEditStart: (selection: EditorSelection) => void;
  onEditChange: (text: string, cursorStart: number, cursorEnd: number) => void;
  onEditEnd: (delta: { deltaRow: number; deltaColumn: number }) => void;
  onEditEscape: () => void;
  onReferenceCycle: (text: string, cursorStart: number, cursorEnd: number) => void;
  html: string;
  cursorStart: number;
  cursorEnd: number;
  focus: boolean;
  isEditing: boolean;
};

export enum FormulaBarTestId {
  CellAddress = 'workbook-formulaBar-cellAddress',
}

const FormulaBar: FunctionComponent<FormulaBarProps> = (properties) => {
  const { isEditing, onEditStart } = properties;

  const formulaBar = useRef<HTMLDivElement>(null);

  const onPointerUp = useCallback(
    (event: PointerEvent) => {
      formulaBar.current?.releasePointerCapture(event.pointerId);
      if (!isEditing) {
        const sel = getSelectedRangeInEditor();
        if (sel) {
          onEditStart(sel);
        }
      }
      event.stopPropagation();
    },
    [isEditing, onEditStart, formulaBar],
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
        <CellBarAddress data-testid={FormulaBarTestId.CellAddress}>
          {properties.cellAddress}
        </CellBarAddress>
        <ChevronContainer>
          <ChevronDownIcon />
        </ChevronContainer>
      </NameContainer>
      <FormulaContainer onPointerUp={onPointerUp} onPointerDown={onPointerDown} ref={formulaBar}>
        <Editor
          onEditChange={properties.onEditChange}
          onEditEnd={properties.onEditEnd}
          onEditEscape={properties.onEditEscape}
          onReferenceCycle={properties.onReferenceCycle}
          display
          html={properties.html}
          mode="edit"
          focus={properties.focus}
          cursorStart={properties.cursorStart}
          cursorEnd={properties.cursorEnd}
        />
      </FormulaContainer>
    </FormulaBarContainer>
  );
};

const CellBarAddress = styled.div`
  width: 100%;
  text-align: center;
`;

// TODO: Implement chevron. - We hide the chevron until implementation is ready
const ChevronContainer = styled.div`
  display: none;
  padding-left: 3px;
`;

const FormulaBarContainer = styled.div`
  flex-shrink: 0;
  display: flex;
  flex-direction: row;
  align-items: center;
  background: ${palette.background.default};
  height: ${formulaBarHeight}px;
`;

const NameContainer = styled.div`
  background: #587af0;
  padding: 5px;
  border-radius: 2px;
  color: #ffffff;
  font-style: normal;
  font-weight: normal;
  font-size: 12px;
  display: flex;
  flex-grow: row;
  min-width: ${headerColumnWidth}px;
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
