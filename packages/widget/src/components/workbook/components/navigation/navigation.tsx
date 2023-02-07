import React, { FunctionComponent } from 'react';
import styled from 'styled-components';
import { PlusIcon, MenuIcon } from 'lucide-react';
import { palette } from 'src/theme';
import * as Menu from 'src/components/uiKit/menu';
import * as Toolbar from 'src/components/uiKit/toolbar';
import { SheetListMenuContent } from './navigationMenus';
import SheetTab from './navigationTab';
import { useWorkbookContext } from '../../workbookContext';

export type NavigationProps = {
  className?: string;
  'data-testid'?: string;
};

const Navigation: FunctionComponent<NavigationProps> = (properties) => {
  const { model, editorState, editorActions } = useWorkbookContext();
  const { selectedSheet } = editorState;

  if (!model) {
    return null;
  }

  const tabs = model.getTabs();
  const onSheetDeleted = (sheet: number) => {
    model.deleteSheet(sheet);
    if (sheet === selectedSheet) {
      editorActions.onSheetSelected(0);
    }
  };

  return (
    <NavigationContainer data-testid={properties['data-testid']}>
      <Toolbar.Root>
        <Menu.Root>
          <Toolbar.Button asChild title="Sheets">
            <Menu.Trigger>
              <MenuIcon size={19} />
            </Menu.Trigger>
          </Toolbar.Button>
          <SheetListMenuContent
            tabs={tabs}
            onSheetSelected={(index: number) => editorActions.onSheetSelected(index)}
          />
        </Menu.Root>
        <Toolbar.Button onClick={() => model.addBlankSheet()} title="Add sheet">
          <PlusIcon size={19} />
        </Toolbar.Button>
      </Toolbar.Root>
      <SheetTabsContainer>
        {tabs.map((tab, index) => {
          const color = tab.color?.RGB;
          const selected = index === selectedSheet;
          return (
            <SheetTab
              key={`${tab.name}-${tab.sheet_id}`}
              name={tab.name}
              selected={selected}
              color={color}
              onSheetSelected={() => editorActions.onSheetSelected(index)}
              onSheetColorChanged={(hex: string) => model.setSheetColor(index, hex)}
              onSheetRenamed={(name: string) => model.renameSheet(index, name)}
              onSheetDeleted={() => onSheetDeleted(index)}
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
