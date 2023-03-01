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
import React, { FunctionComponent, useRef, useState } from 'react';
import * as Menu from 'src/components/uiKit/menu';
import * as Toolbar from 'src/components/uiKit/toolbar';
import styled from 'styled-components';
import { decreaseDecimalPlaces, increaseDecimalPlaces, NumberFormats } from './formatUtil';
import ColorPicker from './colorPicker';
import FormatMenuContent from './formatMenu';
import FormatPicker from './formatPicker';
import { useWorkbookContext } from '../workbookContext';
import Model from '../model';

export type ToolbarProps = {
  className?: string;
  'data-testid'?: string;
};

const TOOLBAR_ICON_SIZE = 19;
const WorkbookToolbar: FunctionComponent<ToolbarProps> = (properties) => {
  const { model, editorState, focusWorkbook } = useWorkbookContext();
  const { selectedSheet, selectedCell, selectedArea } = editorState;
  const [fontColorPickerOpen, setFontColorPickerOpen] = useState(false);
  const [fillColorPickerOpen, setFillColorPickerOpen] = useState(false);
  const [isPickerOpen, setPickerOpen] = useState(false);

  const fontColorButton = useRef(null);
  const fillColorButton = useRef(null);

  if (!model) {
    return null;
  }

  const canEdit = !model.isCellReadOnly(selectedSheet, selectedCell.row, selectedCell.column);
  const cellStyle = model.getCellStyle(selectedSheet, selectedCell.row, selectedCell.column);
  const fontColor = cellStyle.font.color;
  const fillColor = cellStyle.fill.foregroundColor;
  const { numberFormat } = cellStyle;

  const onNumberFormatPicked = (selectedNumberFormat: string) =>
    model.setNumberFormat(selectedSheet, selectedArea, selectedNumberFormat);

  const onToggleAlign = (align: Parameters<Model['toggleAlign']>[2]) =>
    model.toggleAlign(selectedSheet, selectedArea, align);

  const onToggleFontStyle = (fontStyle: Parameters<Model['toggleFontStyle']>[2]) =>
    model.toggleFontStyle(selectedSheet, selectedArea, fontStyle);

  // NB: model.canUndo(), model.canRedo() won't subscribe to changes so if nothing else would
  // change then it wouldn't update the toolbar. But in real usage it can't really happen.
  // We need to rerender if something changes anyway.
  return (
    <ToolbarRoot data-testid={properties['data-testid']}>
      <Toolbar.Button onClick={model.undo} disabled={!model.canUndo()} title="Undo">
        <UndoIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button onClick={model.redo} disabled={!model.canRedo()} title="Redo">
        <RedoIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Separator />
      <Toolbar.Button
        onClick={() => onNumberFormatPicked(NumberFormats.CURRENCY_EUR)}
        disabled={!canEdit}
        title="Format as Euro"
      >
        <EuroIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={() => onNumberFormatPicked(NumberFormats.PERCENTAGE)}
        disabled={!canEdit}
        title="Format as percent"
      >
        <PercentIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={() => onNumberFormatPicked(decreaseDecimalPlaces(numberFormat))}
        disabled={!canEdit}
        title="Decrease decimal places"
      >
        <Icons.DecimalPlacesDecreaseIcon />
      </Toolbar.Button>
      <Toolbar.Button
        onClick={() => onNumberFormatPicked(increaseDecimalPlaces(numberFormat))}
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
          numFmt={numberFormat}
          onChange={(newNumberFormat) => onNumberFormatPicked(newNumberFormat)}
          onExited={(): void => focusWorkbook()}
          setPickerOpen={setPickerOpen}
        />
        <FormatPicker
          numFmt={numberFormat}
          onChange={(newNumberFormat) => onNumberFormatPicked(newNumberFormat)}
          open={isPickerOpen}
          onClose={(): void => setPickerOpen(false)}
          onExited={(): void => focusWorkbook()}
        />
      </Menu.Root>
      <Toolbar.Separator />
      <Toolbar.Button
        $selected={cellStyle.font.bold}
        onClick={() => onToggleFontStyle('bold')}
        disabled={!canEdit}
        title="Bold"
      >
        <BoldIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        $selected={cellStyle.font.italics}
        type="button"
        onClick={() => onToggleFontStyle('italics')}
        disabled={!canEdit}
        title="Italic"
      >
        <ItalicIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        $selected={cellStyle.font.underline}
        onClick={() => onToggleFontStyle('underline')}
        disabled={!canEdit}
        title="Underline"
      >
        <UnderlineIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        $selected={cellStyle.font.strikethrough}
        onClick={() => onToggleFontStyle('strikethrough')}
        disabled={!canEdit}
        title="Strikethrough"
      >
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
      <Toolbar.Button
        $selected={cellStyle.alignment.horizontalAlignment === 'left'}
        onClick={() => onToggleAlign('left')}
        disabled={!canEdit}
        title="Align left"
      >
        <AlignLeftIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        $selected={cellStyle.alignment.horizontalAlignment === 'center'}
        onClick={() => onToggleAlign('center')}
        disabled={!canEdit}
        title="Align center"
      >
        <AlignCenterIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <Toolbar.Button
        $selected={cellStyle.alignment.horizontalAlignment === 'right'}
        onClick={() => onToggleAlign('right')}
        disabled={!canEdit}
        title="Align right"
      >
        <AlignRightIcon size={TOOLBAR_ICON_SIZE} />
      </Toolbar.Button>
      <ColorPicker
        color={fontColor}
        onChange={(color): void => {
          model.setTextColor(selectedSheet, selectedArea, color);
          setFontColorPickerOpen(false);
        }}
        open={fontColorPickerOpen}
      />
      <ColorPicker
        color={fillColor}
        onChange={(color): void => {
          model.setFillColor(selectedSheet, selectedArea, color);
          setFillColorPickerOpen(false);
        }}
        open={fillColorPickerOpen}
      />
    </ToolbarRoot>
  );
};

export default WorkbookToolbar;

const ToolbarRoot = styled(Toolbar.Root)`
  background: transparent;
`;
