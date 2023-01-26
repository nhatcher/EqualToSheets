import styled from 'styled-components';
import StylelessButton from 'src/components/uiKit/button/styleless';
import * as Icons from 'src/components/uiKit/icons';
import React, { FunctionComponent, useEffect, useRef, useState } from 'react';
import { palette } from 'src/theme';
import * as Menu from 'src/components/uiKit/menu';
import { decreaseDecimalPlaces, increaseDecimalPlaces, NumberFormats } from './formatUtil';
import ColorPicker from './colorPicker';
import FormatMenuContent from './formatMenu';
import { Area } from '../util';

const toolbarHeight = 40;

export type ToolbarProps = {
  className?: string;
  'data-testid'?: string;
  canUndo: boolean;
  canRedo: boolean;
  onRedo: () => void;
  onUndo: () => void;
  onToggleUnderline: () => void;
  onToggleBold: () => void;
  onToggleItalic: () => void;
  onToggleStrike: () => void;
  onToggleAlignLeft: () => void;
  onToggleAlignCenter: () => void;
  onToggleAlignRight: () => void;
  onTextColorPicked: (hex: string) => void;
  onFillColorPicked: (hex: string) => void;
  onNumberFormatPicked: (numberFmt: string) => void;
  focusWorkbook: () => void;
  fillColor: string;
  fontColor: string;
  bold: boolean;
  underline: boolean;
  italic: boolean;
  strike: boolean;
  alignment: string;
  canEdit: boolean;
  numFmt: string;
  selectedArea: Area;
};

const Toolbar: FunctionComponent<ToolbarProps> = (properties) => {
  const [fontColor, setFontColor] = useState<string>(properties.fontColor);
  const [fillColor, setFillColor] = useState<string>(properties.fillColor);
  const [fontColorPickerOpen, setFontColorPickerOpen] = useState(false);
  const [fillColorPickerOpen, setFillColorPickerOpen] = useState(false);

  const fontColorButton = useRef(null);
  const fillColorButton = useRef(null);

  useEffect(() => {
    setFillColor(properties.fillColor);
    setFontColor(properties.fontColor);
  }, [properties.fillColor, properties.fontColor]);

  const { canEdit } = properties;

  return (
    <ToolbarContainer data-testid={properties['data-testid']}>
      <Button
        type="button"
        $pressed={false}
        onClick={properties.onUndo}
        disabled={!properties.canUndo}
        title="workbook.toolbar.undo_button_title"
      >
        <Icons.UndoIcon />
      </Button>
      <Button
        type="button"
        $pressed={false}
        onClick={properties.onRedo}
        disabled={!properties.canRedo}
        title="workbook.toolbar.redo_button_title"
      >
        <Icons.RedoIcon />
      </Button>
      <Divider />
      <Button
        type="button"
        $pressed={false}
        onClick={(): void => {
          properties.onNumberFormatPicked(NumberFormats.CURRENCY_EUR);
        }}
        disabled={!canEdit}
        title="workbook.toolbar.euro_button_title"
      >
        <Icons.EuroIcon />
      </Button>
      <Button
        type="button"
        $pressed={false}
        onClick={(): void => {
          properties.onNumberFormatPicked(NumberFormats.PERCENTAGE);
        }}
        disabled={!canEdit}
        title="workbook.toolbar.percentage_button_title"
      >
        <Icons.PercentIcon />
      </Button>
      <Button
        type="button"
        $pressed={false}
        onClick={(): void => {
          properties.onNumberFormatPicked(decreaseDecimalPlaces(properties.numFmt));
        }}
        disabled={!canEdit}
        title="workbook.toolbar.decimal_places_decrease_button_title"
      >
        <Icons.DecimalPlacesDecreaseIcon />
      </Button>
      <Button
        type="button"
        $pressed={false}
        onClick={(): void => {
          properties.onNumberFormatPicked(increaseDecimalPlaces(properties.numFmt));
        }}
        disabled={!canEdit}
        title="workbook.toolbar.decimal_places_increase_button_title"
      >
        <Icons.DecimalPlacesIncreaseIcon />
      </Button>
      <Menu.Root>
        <Menu.Trigger disabled={!canEdit} title="workbook.toolbar.format_button_title">
          {'123'}
          <Icons.ChevronDownIcon />
        </Menu.Trigger>
        <FormatMenuContent
          numFmt={properties.numFmt}
          onChange={(numberFmt): void => {
            properties.onNumberFormatPicked(numberFmt);
          }}
          onExited={(): void => properties.focusWorkbook()}
        />
      </Menu.Root>
      <Divider />
      <Button
        type="button"
        $pressed={properties.bold}
        onClick={properties.onToggleBold}
        disabled={!canEdit}
        title="workbook.toolbar.bold_button_title"
      >
        <Icons.BoldIcon />
      </Button>
      <Button
        type="button"
        $pressed={properties.italic}
        onClick={properties.onToggleItalic}
        disabled={!canEdit}
        title="workbook.toolbar.italic_button_title"
      >
        <Icons.ItalicIcon />
      </Button>
      <Button
        type="button"
        $pressed={properties.underline}
        onClick={properties.onToggleUnderline}
        disabled={!canEdit}
        title="workbook.toolbar.underline_button_title"
      >
        <Icons.UnderlineIcon />
      </Button>
      <Button
        type="button"
        $pressed={properties.strike}
        onClick={properties.onToggleStrike}
        disabled={!canEdit}
        title="workbook.toolbar.strike_button_title"
      >
        <Icons.StrikethroughIcon />
      </Button>
      <Divider />
      <Button
        type="button"
        $pressed={false}
        disabled={!canEdit}
        title="workbook.toolbar.font_color_button_title"
        ref={fontColorButton}
        $underlinedColor={fontColor}
        onClick={() => setFontColorPickerOpen(true)}
      >
        <Icons.TypeIcon />
      </Button>
      <Button
        type="button"
        $pressed={false}
        disabled={!canEdit}
        title="workbook.toolbar.fill_button_title"
        ref={fillColorButton}
        $underlinedColor={fillColor}
        onClick={() => setFillColorPickerOpen(true)}
      >
        <Icons.PainBucketIcon />
      </Button>

      <Divider />
      <Button
        type="button"
        $pressed={properties.alignment === 'left'}
        onClick={properties.onToggleAlignLeft}
        disabled={!canEdit}
        title="workbook.toolbar.align_left_button_title"
      >
        <Icons.AlignLeftIcon />
      </Button>
      <Button
        type="button"
        $pressed={properties.alignment === 'center'}
        onClick={properties.onToggleAlignCenter}
        disabled={!canEdit}
        title="workbook.toolbar.align_center_button_title"
      >
        <Icons.AlignCenterIcon />
      </Button>
      <Button
        type="button"
        $pressed={properties.alignment === 'right'}
        onClick={properties.onToggleAlignRight}
        disabled={!canEdit}
        title="workbook.toolbar.align_right_button_title"
      >
        <Icons.AlignRightIcon />
      </Button>
      <ColorPicker
        color={fontColor}
        onChange={(color): void => {
          properties.onTextColorPicked(color);
          setFontColorPickerOpen(false);
        }}
        open={fontColorPickerOpen}
      />
      <ColorPicker
        color={fillColor}
        onChange={(color): void => {
          properties.onFillColorPicked(color);
          setFillColorPickerOpen(false);
        }}
        open={fillColorPickerOpen}
      />
    </ToolbarContainer>
  );
};

const Divider = styled.div`
  display: inline-flex;
  height: 10px;
  width: 1px;
  border-left: 1px solid #d3d6e9;
  margin-left: 5px;
  margin-right: 5px;
`;

const ToolbarContainer = styled.div`
  display: flex;
  flex-shrink: 0;
  flex-grow: row;
  align-items: center;
  background: ${palette.background.default};
  height: ${toolbarHeight}px;
  line-height: ${toolbarHeight}px;
  border-bottom: 1px solid ${palette.grays.gray2};
`;

type TypeButtonProperties = { $pressed: boolean; $underlinedColor?: string };
const Button = styled(StylelessButton).attrs<TypeButtonProperties>((properties) => ({
  'aria-pressed': properties.$pressed,
}))<TypeButtonProperties>`
  width: 23px;
  height: 23px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border-radius: 2px;
  margin-right: 5px;
  transition: all 0.2s;

  ${({ disabled, $pressed, $underlinedColor }): string => {
    if (disabled) {
      return `
      color: ${palette.grays.gray3};
      cursor: default;
    `;
    }
    return `
      border-top: ${$underlinedColor ? '3px solid #FFF' : 'none'};
      border-bottom: ${$underlinedColor ? `3px solid ${$underlinedColor}` : 'none'};
      color: ${palette.text.primary};
      background-color: ${$pressed ? palette.grays.gray3 : '#FFF'};
      &:hover {
        background-color: ${palette.grays.gray2};
        border-top-color: ${palette.grays.gray2};
      }
    `;
  }}
`;

export default Toolbar;
