import styled from 'styled-components';
import React, { FunctionComponent } from 'react';
import { useWorkbookContext } from '../workbookContext';
import { getCellAddress } from '../util';

const formulaBarHeight = 30;

type FormulaBarProps = {
  className?: string;
  'data-testid'?: string;
};

export enum FormulaBarTestId {
  CellAddress = 'workbook-formulaBar-cellAddress',
}

const FormulaBar: FunctionComponent<FormulaBarProps> = (properties) => {
  const { editorState, formulaBarEditor } = useWorkbookContext();
  const { selectedArea, selectedCell } = editorState;
  const cellAddress = getCellAddress(selectedArea, selectedCell);

  return (
    <FormulaBarContainer data-testid={properties['data-testid']}>
      <NameContainer>
        <CellBarAddress data-testid={FormulaBarTestId.CellAddress}>{cellAddress}</CellBarAddress>
      </NameContainer>
      <FormulaContainer ref={formulaBarEditor} />
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
  line-height: 100%;
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
