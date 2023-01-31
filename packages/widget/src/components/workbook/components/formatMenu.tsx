import React, { FunctionComponent, ReactNode } from 'react';
import * as Menu from 'src/components/uiKit/menu';
import styled from 'styled-components';
import { palette } from 'src/theme';
import { NumberFormats } from './formatUtil';

type FormatMenuProps = {
  numFmt: string;
  onChange: (numberFmt: string) => void;
  setPickerOpen: (open: boolean) => void;
  onExited?: () => void;
  children?: ReactNode;
};

const FormatMenuContent: FunctionComponent<FormatMenuProps> = (properties) => {
  const { onChange, setPickerOpen } = properties;

  return (
    <Menu.Content onExited={properties.onExited}>
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.AUTO)}>
        {'Auto'}
      </MenuItemWrapped>
      {/** TODO: Text option that transforms into plain text */}
      <Menu.Divider />
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.NUMBER)}>
        <MenuItemText>{'Number'}</MenuItemText>
        <MenuItemText>{'1,000.00'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.PERCENTAGE)}>
        <MenuItemText>{'Percentage'}</MenuItemText>
        <MenuItemText>{'10%'}</MenuItemText>
      </MenuItemWrapped>

      <Menu.Divider />
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.CURRENCY_EUR)}>
        <MenuItemText>{'Euro (EUR)'}</MenuItemText>
        <MenuItemText>{'€'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.CURRENCY_USD)}>
        <MenuItemText>{'Dollar (USD)'}</MenuItemText>
        <MenuItemText>{'$'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.CURRENCY_GBP)}>
        <MenuItemText>{'British Pound (GBP)'}</MenuItemText>
        <MenuItemText>{'£'}</MenuItemText>
      </MenuItemWrapped>

      <Menu.Divider />
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.DATE_SHORT)}>
        <MenuItemText>{'Short date'}</MenuItemText>
        <MenuItemText>{'03/03/2021'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onSelect={(): void => onChange(NumberFormats.DATE_LONG)}>
        <MenuItemText>{'Long date'}</MenuItemText>
        <MenuItemText>{'Wednesday, March 3, 2021'}</MenuItemText>
      </MenuItemWrapped>

      <Menu.Divider />
      <MenuItemWrapped onSelect={(): void => setPickerOpen(true)}>{'Custom'}</MenuItemWrapped>
    </Menu.Content>
  );
};

export default FormatMenuContent;

const MenuItemWrapped = styled(Menu.Item)`
  display: flex;
  justify-content: space-between;
`;

const MenuItemText = styled.div`
  &:last-child {
    color: ${palette.text.secondary};
  }
  & + & {
    margin-left: 20px;
  }
`;
