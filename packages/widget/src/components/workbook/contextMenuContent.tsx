import React, { FunctionComponent } from 'react';
import * as Menu from 'src/components/uiKit/menu';
import styled from 'styled-components';
import { fonts } from 'src/theme';

export const RowContextMenuContent: FunctionComponent<{
  isMenuOpen: boolean;
  row: number;
  onInsertRow: (insertBeforeRow: number) => void;
  onDeleteRow: (row: number) => void;
  onClose: () => void;
  anchorEl: HTMLDivElement | null;
}> = (properties) => {
  const { onDeleteRow, onInsertRow, row } = properties;

  return (
    <StyledMenuContent>
      <StyledMenuItem
        onClick={(): void => {
          onInsertRow(row);
        }}
      >
        {'Insert row above'}
      </StyledMenuItem>
      <StyledMenuItem
        onClick={(): void => {
          onInsertRow(row + 1);
        }}
      >
        {'Insert row below'}
      </StyledMenuItem>
      <StyledMenuDivider />
      <StyledMenuItem
        onClick={(): void => {
          onDeleteRow(row);
        }}
      >
        {'Delete row'}
      </StyledMenuItem>
    </StyledMenuContent>
  );
};

export const ColumnContextMenuContent: FunctionComponent<{
  isMenuOpen: boolean;
  column: number;
  onInsertColumn: (insertBeforeColumn: number) => void;
  onDeleteColumn: (column: number) => void;
  onClose: () => void;
  anchorEl: HTMLDivElement | null;
}> = (properties) => {
  const { onDeleteColumn, onInsertColumn, column } = properties;

  return (
    <StyledMenuContent>
      <StyledMenuItem
        onClick={(): void => {
          onInsertColumn(column);
        }}
      >
        {'Insert column before'}
      </StyledMenuItem>
      <StyledMenuItem
        onClick={(): void => {
          onInsertColumn(column + 1);
        }}
      >
        {'Insert column after'}
      </StyledMenuItem>
      <StyledMenuDivider />
      <StyledMenuItem
        onClick={(): void => {
          onDeleteColumn(column);
        }}
      >
        {'Delete column'}
      </StyledMenuItem>
    </StyledMenuContent>
  );
};

const StyledMenuItem = styled(Menu.Item)`
  font-size: 14px;
  padding: 10px 20px;
`;

const StyledMenuDivider = styled(Menu.Divider)`
  margin: 0px 20px;
`;

const StyledMenuContent = styled(Menu.Content)`
  border-radius: 10px;
  font-family: ${fonts.regular};
`;
