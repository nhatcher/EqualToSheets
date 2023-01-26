import React, { FunctionComponent } from 'react';
import styled from 'styled-components';
import * as Menu from 'src/components/uiKit/menu';
import { TabsInput } from './common';

type SheetListMenuProps = {
  onSheetSelected: (index: number) => void;
  tabs: TabsInput[];
};

const SheetListItem = styled.div`
  display: flex;
  align-items: center;
`;

export const SheetListMenuContent: FunctionComponent<SheetListMenuProps> = (properties) => {
  const { onSheetSelected, tabs } = properties;

  return (
    <Menu.Content>
      {tabs.map((tab, index) => (
        <Menu.Item
          key={`sheet-list-item-${tab.name}-${tab.sheet_id}`}
          onClick={(): void => onSheetSelected(index)}
        >
          <SheetListItem>
            <SheetListItemColor $color={tab.color?.RGB} />
            <SheetListItemName>{tab.name}</SheetListItemName>
          </SheetListItem>
        </Menu.Item>
      ))}
    </Menu.Content>
  );
};

const SheetListItemColor = styled.div<{ $color?: string }>`
  width: 10px;
  height: 10px;
  border-radius: 2px;
  margin-right: 10px;
  margin-bottom: 2px; // TODO [MVP]: Is there better way to align it vertically with the text
  background-color: ${({ $color }): string => $color ?? '#fff'};
`;

const SheetListItemName = styled.span``;
