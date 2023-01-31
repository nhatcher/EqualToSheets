import React, { FunctionComponent } from 'react';
import styled from 'styled-components';
import { PlusIcon, MenuIcon } from 'lucide-react';
import { palette } from 'src/theme';
import * as Menu from 'src/components/uiKit/menu';
import * as Toolbar from 'src/components/uiKit/toolbar';
import { SheetListMenuContent } from './navigationMenus';
import SheetTab from './navigationTab';
import { TabsInput } from './common';

type TabOption = {
  hideRename?: boolean;
  hideDelete?: boolean;
};

export type NavigationProps = {
  className?: string;
  'data-testid'?: string;
  tabs: TabsInput[];
  tabsOptions?: Record<TabsInput['sheet_id'], TabOption>;
  selectedSheet: number;
  onSheetSelected: (index: number) => void;
  onAddBlankSheet: () => void;
  onSheetColorChanged: (hex: string) => void;
  onSheetRenamed: (name: string) => void;
  onSheetDeleted: () => void;
  readOnly?: boolean;
  disabled?: boolean;
};

const Navigation: FunctionComponent<NavigationProps> = (properties) => {
  const {
    tabs,
    tabsOptions,
    selectedSheet,
    onSheetSelected,
    onAddBlankSheet,
    onSheetColorChanged,
    onSheetRenamed,
    onSheetDeleted,
    readOnly,
  } = properties;

  return (
    <NavigationContainer data-testid={properties['data-testid']}>
      {!readOnly && (
        <Toolbar.Root>
          <Menu.Root>
            <Toolbar.Button asChild disabled={properties.disabled} title="Sheets">
              <Menu.Trigger>
                <MenuIcon size={19} />
              </Menu.Trigger>
            </Toolbar.Button>
            <SheetListMenuContent tabs={tabs} onSheetSelected={onSheetSelected} />
          </Menu.Root>
          <Menu.Root>
            <Toolbar.Button asChild disabled={properties.disabled} title="Add sheet">
              <Menu.Trigger>
                <PlusIcon size={19} />
              </Menu.Trigger>
            </Toolbar.Button>
            <Menu.Content side="top">
              <Menu.Item onSelect={onAddBlankSheet}>{'Add blank sheet'}</Menu.Item>
            </Menu.Content>
          </Menu.Root>
        </Toolbar.Root>
      )}
      <SheetTabsContainer>
        {tabs.map((tab, index) => {
          const color = tab.color?.RGB;
          const selected = index === selectedSheet;
          const tabOptions = tabsOptions?.[tab.sheet_id];
          return (
            <SheetTab
              key={`${tab.name}-${tab.sheet_id}`}
              name={tab.name}
              selected={selected}
              color={color}
              onSheetSelected={(): void => onSheetSelected(index)}
              onSheetColorChanged={onSheetColorChanged}
              onSheetRenamed={onSheetRenamed}
              onSheetDeleted={onSheetDeleted}
              readOnly={readOnly}
              hideDelete={tabOptions?.hideDelete}
              hideRename={tabOptions?.hideRename}
            />
          );
        })}
      </SheetTabsContainer>
    </NavigationContainer>
  );
};

const NavigationContainer = styled.div`
  flex-shrink: 0;
  display: flex;
  flex-direction: row;
  align-items: center;
  height: 40px;
  line-height: 40px;
  padding-left: 28px;
  background: ${palette.background.default};
  color: ${palette.text.primary};
  overflow: hidden;
`;

const SheetTabsContainer = styled.div`
  display: flex;
  flex-direction: row;
  text-align: left;
  cursor: pointer;
  margin-left: 10px;
  &:first-child {
    margin-left: 0px;
  }
`;

export default Navigation;
