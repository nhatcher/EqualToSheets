import {
  UndoIcon,
  RedoIcon,
  EuroIcon,
  PercentIcon,
  ChevronDownIcon,
  BoldIcon,
  ItalicIcon,
  UnderlineIcon,
  StrikethroughIcon,
  TypeIcon,
  PaintBucketIcon,
  AlignLeftIcon,
  AlignCenterIcon,
  AlignRightIcon,
} from 'lucide-react';
import * as Icons from 'src/components/uiKit/icons';
import React, { FunctionComponent, useEffect, useRef, useState } from 'react';
import * as Menu from 'src/components/uiKit/menu';
import * as Toolbar from 'src/components/uiKit/toolbar';
import { decreaseDecimalPlaces, increaseDecimalPlaces, NumberFormats } from './formatUtil';
import ColorPicker from './colorPicker';
import FormatMenuContent from './formatMenu';
import FormatPicker from './formatPicker';
import { Area } from '../util';

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

const TOOLBAR_ICON_SIZE = 19;
const WorkbookToolbar: FunctionComponent<ToolbarProps> = (properties) => {
  const [fontColor, setFontColor] = useState<string>(properties.fontColor);
  const [fillColor, setFillColor] = useState<string>(properties.fillColor);
  const [fontColorPickerOpen, setFontColorPickerOpen] = useState(false);
  const [fillColorPickerOpen, setFillColorPickerOpen] = useState(false);
  const [isPickerOpen, setPickerOpen] = useState(false);

  const fontColorButton = useRef(null);
  const fillColorButton = useRef(null);

  useEffect(() => {
    setFillColor(properties.fillColor);
    setFontColor(properties.fontColor);
  }, [properties.fillColor, properties.fontColor]);

  const { canEdit } = properties;

  return (
    <Toolbar.Root data-testid={properties['data-testid']}>
      <Toolbar.Button onClick={properties.onUndo} disabled={!properties.canUndo} title="Undo">
        <UndoIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button onClick={properties.onRedo} disabled={!properties.canRedo} title="Redo">
        <RedoIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Separator />
      <Toolbar.Button
        onClick={(): void => {
          properties.onNumberFormatPicked(NumberFormats.CURRENCY_EUR);
        }}
        disabled={!canEdit}
        title="Format as Euro"
      >
        <EuroIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={(): void => {
          properties.onNumberFormatPicked(NumberFormats.PERCENTAGE);
        }}
        disabled={!canEdit}
        title="Format as percent"
      >
        <PercentIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={(): void => {
          properties.onNumberFormatPicked(decreaseDecimalPlaces(properties.numFmt));
        }}
        disabled={!canEdit}
        title="Decrease decimal places"
      >
        <Icons.DecimalPlacesDecreaseIcon />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={(): void => {
          properties.onNumberFormatPicked(increaseDecimalPlaces(properties.numFmt));
        }}
        disabled={!canEdit}
        title="Increase decimal places"
      >
        <Icons.DecimalPlacesIncreaseIcon />
      </Toolbar.Button>
      <Menu.Root>
        <Toolbar.Button asChild style={{ width: 30 }}>
          <Menu.Trigger disabled={!canEdit} title="workbook.toolbar.format_button_title">
            {'123'}
            <ChevronDownIcon size={14} />
          </Menu.Trigger>
        </Toolbar.Button>
        <FormatMenuContent
          numFmt={properties.numFmt}
          onChange={(numberFmt): void => {
            properties.onNumberFormatPicked(numberFmt);
          }}
          onExited={(): void => properties.focusWorkbook()}
          setPickerOpen={setPickerOpen}
        />
        <FormatPicker
          numFmt={properties.numFmt}
          onChange={(numberFmt): void => {
            properties.onNumberFormatPicked(numberFmt);
          }}
          open={isPickerOpen}
          onClose={(): void => setPickerOpen(false)}
          onExited={(): void => properties.focusWorkbook()}
        />
      </Menu.Root>
      <Toolbar.Separator />
      <Toolbar.Button onClick={properties.onToggleBold} disabled={!canEdit} title="Bold">
        <BoldIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        type="button"
        onClick={properties.onToggleItalic}
        disabled={!canEdit}
        title="Italic"
      >
        <ItalicIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button onClick={properties.onToggleUnderline} disabled={!canEdit} title="Underline">
        <UnderlineIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button onClick={properties.onToggleStrike} disabled={!canEdit} title="Strikethrough">
        <StrikethroughIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Separator />
      <Toolbar.Button
        disabled={!canEdit}
        title="Font color"
        ref={fontColorButton}
        $underlinedColor={fontColor}
        onClick={() => setFontColorPickerOpen(true)}
      >
        <TypeIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        disabled={!canEdit}
        title="Fill color"
        ref={fillColorButton}
        $underlinedColor={fillColor}
        onClick={() => setFillColorPickerOpen(true)}
      >
        <PaintBucketIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>

      <Toolbar.Separator />
      <Toolbar.Button onClick={properties.onToggleAlignLeft} disabled={!canEdit} title="Align left">
        <AlignLeftIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={properties.onToggleAlignCenter}
        disabled={!canEdit}
        title="Align center"
      >
        <AlignCenterIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={properties.onToggleAlignRight}
        disabled={!canEdit}
        title="Align right"
      >
        <AlignRightIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
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
    </Toolbar.Root>
  );
};

export default WorkbookToolbar;
