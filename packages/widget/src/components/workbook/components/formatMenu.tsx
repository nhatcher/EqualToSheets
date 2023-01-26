import React, { FunctionComponent, useState, ReactNode } from 'react';
import * as Menu from 'src/components/uiKit/menu';
import styled from 'styled-components';
import { palette } from 'src/theme';
import { NumberFormats } from './formatUtil';
import FormatPicker from './formatPicker';

type FormatMenuProps = {
  numFmt: string;
  onChange: (numberFmt: string) => void;
  onExited?: () => void;
  children?: ReactNode;
};

const FormatMenuContent: FunctionComponent<FormatMenuProps> = (properties) => {
  const { onChange } = properties;
  const [isPickerOpen, setPickerOpen] = useState(false);

  return (
    <Menu.Content onExited={properties.onExited}>
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.AUTO)}>
        {'workbook.toolbar.format_menu.auto'}
      </MenuItemWrapped>
      {/** TODO: Text option that transforms into plain text */}
      <Menu.Divider />
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.NUMBER)}>
        <MenuItemText>{'workbook.toolbar.format_menu.number'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.number_example'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.PERCENTAGE)}>
        <MenuItemText>{'workbook.toolbar.format_menu.percentage'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.percentage_example'}</MenuItemText>
      </MenuItemWrapped>

      <Menu.Divider />
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.CURRENCY_EUR)}>
        <MenuItemText>{'workbook.toolbar.format_menu.currency_eur'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.currency_eur_example'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.CURRENCY_USD)}>
        <MenuItemText>{'workbook.toolbar.format_menu.currency_usd'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.currency_usd_example'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.CURRENCY_GBP)}>
        <MenuItemText>{'workbook.toolbar.format_menu.currency_gbp'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.currency_gbp_example'}</MenuItemText>
      </MenuItemWrapped>

      <Menu.Divider />
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.DATE_SHORT)}>
        <MenuItemText>{'workbook.toolbar.format_menu.date_short'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.date_short_example'}</MenuItemText>
      </MenuItemWrapped>
      <MenuItemWrapped onClick={(): void => onChange(NumberFormats.DATE_LONG)}>
        <MenuItemText>{'workbook.toolbar.format_menu.date_long'}</MenuItemText>
        <MenuItemText>{'workbook.toolbar.format_menu.date_long_example'}</MenuItemText>
      </MenuItemWrapped>

      <Menu.Divider />
      <MenuItemWrapped onClick={(): void => setPickerOpen(true)}>{'Custom'}</MenuItemWrapped>
      <FormatPicker
        numFmt={properties.numFmt}
        onChange={properties.onChange}
        open={isPickerOpen}
        onClose={(): void => setPickerOpen(false)}
        onExited={properties.onExited}
      />
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
