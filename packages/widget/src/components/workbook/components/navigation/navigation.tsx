import React, { FunctionComponent } from 'react';
import styled from 'styled-components';
import StylelessButton from 'src/components/uiKit/button/styleless';
import { PlusIcon, MenuIcon } from 'src/components/uiKit/icons';
import { palette } from 'src/theme';
import * as Menu from 'src/components/uiKit/menu';
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
        <>
          <Menu.Root>
            <Menu.Trigger disabled={properties.disabled}>
              <MenuIcon />
            </Menu.Trigger>
            <SheetListMenuContent tabs={tabs} onSheetSelected={onSheetSelected} />
          </Menu.Root>
          <Menu.Root>
            <Menu.Trigger>
              <NavigationButton
                disabled={properties.disabled}
                title="workbook.navigation.add_sheet_button_title"
              >
                <PlusIcon />
              </NavigationButton>
            </Menu.Trigger>
            <Menu.Content>
              <Menu.Item onClick={onAddBlankSheet} disabled>
                {'Add blank sheet'}
              </Menu.Item>
            </Menu.Content>
          </Menu.Root>
        </>
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

const NavigationButton = styled(StylelessButton)`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 20px;
  height: 20px;
  border-radius: 2px;
  transition-duration: 0.2s;
  transition-property: background-color, color;

  color: ${({ disabled }): string => (!disabled ? palette.text.primary : palette.grays.gray2)};

  ${({ disabled }): string =>
    !disabled
      ? `
    &:hover {
      background: ${palette.grays.gray2};
    }
    &:active {
      background: ${palette.grays.gray3};
    }
    `
      : ``}

  & ~ & {
    margin-left: 5px;
  }
`;

export default Navigation;
