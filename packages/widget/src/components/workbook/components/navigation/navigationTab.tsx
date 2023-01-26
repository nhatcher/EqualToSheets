import React, { FunctionComponent, useRef, useState } from 'react';
import styled from 'styled-components';
import ReserveSpaceForFontWeight from 'src/components/uiKit/util/reserveSpaceForFontWeight';
import StylelessButton from 'src/components/uiKit/button/styleless';
import * as Menu from 'src/components/uiKit/menu';
import { ChevronDownIcon } from 'src/components/uiKit/icons';
import PromptDialog, { PromptDialogSubmitResult } from 'src/components/uiKit/dialog/prompt';
import { palette } from 'src/theme';
import ColorPicker from '../colorPicker';

export enum NavigationTabTestId {
  Container = 'workbook-navigation-navigationTab-container',
  TabColor = 'workbook-navigation-navigationTab-tabColor',
}

type SheetTabProps = {
  className?: string;
  name: string;
  selected: boolean;
  color?: string;
  onSheetSelected: () => void;
  onSheetColorChanged: (hex: string) => void;
  onSheetRenamed: (name: string) => void;
  onSheetDeleted: () => void;
  readOnly?: boolean;
  hideDelete?: boolean;
  hideRename?: boolean;
};

const SheetTab: FunctionComponent<SheetTabProps> = (properties) => {
  const { color, name, selected, onSheetSelected, readOnly, hideRename, hideDelete } = properties;

  const [renameDialogOpen, setRenameDialogOpen] = useState(false);
  const [sheetColor, setSheetColor] = useState(color);
  const [displayPicker, setDisplayPicker] = useState(false);

  const sheetTabReference = useRef<HTMLButtonElement | null>(null);

  return (
    <SheetTabContainer
      data-testid={NavigationTabTestId.Container}
      $selected={selected}
      onClick={onSheetSelected}
    >
      <SpaceContainer />
      <NameWrapper>
        <NameButton ref={sheetTabReference}>
          <ReserveSpaceForFontWeight maxFontWeight="600">{name}</ReserveSpaceForFontWeight>
        </NameButton>
        {!readOnly && (
          <Menu.Root>
            <Menu.Trigger title="workbook.navigation.sheet_options_button_title">
              <ChevronDownIcon />
            </Menu.Trigger>
            <Menu.Content>
              {!hideRename && (
                <Menu.Item
                  onClick={(): void => {
                    setRenameDialogOpen(true);
                  }}
                >
                  {'workbook.navigation.rename_sheet'}
                </Menu.Item>
              )}
              <Menu.Item onClick={() => setDisplayPicker(true)}>
                <span>{'workbook.navigation.change_sheet_color'}</span>
              </Menu.Item>
              {!hideDelete && (
                <Menu.Item
                  onClick={(): void => {
                    properties.onSheetDeleted();
                  }}
                >
                  {'workbook.navigation.delete_sheet'}
                </Menu.Item>
              )}
            </Menu.Content>
          </Menu.Root>
        )}
      </NameWrapper>
      <PromptDialog
        open={renameDialogOpen}
        onClose={() => {
          setRenameDialogOpen(false);
        }}
        onSubmit={(newName: string): Promise<PromptDialogSubmitResult> => {
          // Regrettably we cannot get an error from here
          properties.onSheetRenamed(newName);
          return Promise.resolve({ success: true });
        }}
        title="workbook.navigation.menu_rename_sheet"
        label="workbook.navigation.menu_new_name"
        defaultValue={name}
      />
      <ColorContainer
        data-testid={NavigationTabTestId.TabColor}
        $color={color}
        $selected={selected}
      />
      <ColorPicker
        color={sheetColor || ''}
        onChange={(new_color): void => {
          properties.onSheetColorChanged(new_color);
          setSheetColor(new_color);
          setDisplayPicker(false);
        }}
        open={displayPicker}
      />
    </SheetTabContainer>
  );
};

export default SheetTab;

const SheetTabContainer = styled.div<{ $selected: boolean }>`
  flex-shrink: 0;
  display: flex;
  flex-direction: column;
  letter-spacing: 0em;
  text-align: left;
  margin-right: 20px;
  cursor: pointer;
  transition: all 0.1s;
  font-weight: ${({ $selected }): number => ($selected ? 600 : 400)};
  color: ${({ $selected }): string => ($selected ? palette.text.primary : palette.grays.gray4)};

  &:hover {
    color: ${palette.text.primary};
  }
`;

const SpaceContainer = styled.div`
  flex-basis: 3px;
  width: 100%;
`;

const NameWrapper = styled.div`
  flex-shrink: 0;
  display: flex;
  flex-direction: row;
  flex-basis: 34px;
  line-height: 34px;
`;

const ColorContainer = styled.div<{ $color?: string; $selected: boolean }>`
  background-color: ${({ $color }): string => $color ?? '#fff'};

  flex-basis: 3px;
  width: 100%;
  border-radius: 3px 3px 0px 0px;
`;

const NameButton = styled(StylelessButton)`
  flex-shrink: 0;
  padding-right: 4px;
  font-family: inherit;
  color: inherit;
  font-weight: inherit;
`;
