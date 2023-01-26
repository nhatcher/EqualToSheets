import styled from 'styled-components';
import React, { FunctionComponent, useEffect, useRef, useState } from 'react';
import { HexColorInput, HexColorPicker } from 'react-colorful';
import StylelessButton from 'src/components/uiKit/button/styleless';
import Dialog from 'src/components/uiKit/dialog';
import { fonts, palette } from 'src/theme';

type ColorPickerProps = {
  className?: string;
  color: string;
  onChange: (color: string) => void;
  open: boolean;
};

const colorPickerWidth = 240;
const colorPickerPadding = 15;
const colorfulHeight = 185; // 150 + 15 + 20

// TODO: We want to use a popover
const ColorPicker: FunctionComponent<ColorPickerProps> = (properties) => {
  const [color, setColor] = useState(properties.color);
  const recentColors = useRef<string[]>([]);

  const closePicker = (newColor: string): void => {
    const maxRecentColors = 14;
    properties.onChange(newColor);
    const colors = recentColors.current.filter((c) => c !== newColor);
    recentColors.current = [newColor, ...colors].slice(0, maxRecentColors);
  };

  useEffect(() => {
    setColor(properties.color);
  }, [properties.color]);

  const presetColors = [
    '#FFFFFF',
    '#1B717E',
    '#59B9BC',
    '#3BB68A',
    '#8CB354',
    '#F8CD3C',
    '#EC5753',
    '#A23C52',
    '#D03627',
    '#523E93',
    '#3358B7',
  ];

  return (
    <Dialog
      title="Select the color"
      open={properties.open}
      onClose={(): void => closePicker(color)}
    >
      <ColorPickerDialog>
        <HexColorPicker
          color={color}
          onChange={(newColor): void => {
            setColor(newColor);
          }}
        />
        <ColorPickerInput>
          <HexWrapper>
            <HexLabel>{'Hex'}</HexLabel>
            <HexColorInputBox>
              <HashLabel>{'#'}</HashLabel>
              <HexColorInput
                color={color}
                onChange={(newColor): void => {
                  setColor(newColor);
                }}
              />
            </HexColorInputBox>
          </HexWrapper>
          <Swatch $color={color} />
        </ColorPickerInput>
        <HorizontalDivider />
        <ColorList>
          {presetColors.map((presetColor) => (
            <Button
              key={presetColor}
              $color={presetColor}
              onClick={(): void => {
                closePicker(presetColor);
              }}
            />
          ))}
        </ColorList>
        <HorizontalDivider />
        <RecentLabel>{'Recent'}</RecentLabel>
        <ColorList>
          {recentColors.current.map((recentColor) => (
            <Button
              key={recentColor}
              $color={recentColor}
              onClick={(): void => {
                closePicker(recentColor);
              }}
            />
          ))}
        </ColorList>
      </ColorPickerDialog>
    </Dialog>
  );
};

const RecentLabel = styled.div`
  font-size: 12px;
  color: ${palette.text.secondary};
`;

const ColorList = styled.div`
  display: flex;
  flex-wrap: wrap;
  flex-direction: row;
`;

const Button = styled(StylelessButton)<{ $color: string }>`
  width: 20px;
  height: 20px;
  ${({ $color }): string => {
    if ($color.toUpperCase() === '#FFFFFF') {
      return `border: 1px solid ${palette.grays.gray3};`;
    }
    return `border: 1px solid ${$color};`;
  }}
  background-color: ${({ $color }): string => $color};
  box-sizing: border-box;
  margin-top: 10px;
  margin-right: 10px;
  border-radius: 2px;
`;

const HorizontalDivider = styled.div`
  height: 0px;
  width: 100%;
  border-top: 1px solid ${palette.grays.gray2};
  margin-top: 15px;
  margin-bottom: 5px;
`;

const ColorPickerDialog = styled.div`
  background: ${palette.background.default};
  width: ${colorPickerWidth}px;
  padding: 15px;
  display: flex;
  flex-direction: column;

  & .react-colorful {
    height: ${colorfulHeight}px;
    width: ${colorPickerWidth - colorPickerPadding * 2}px;
  }
  & .react-colorful__saturation {
    border-bottom: none;
    border-radius: 5px;
  }
  & .react-colorful__hue {
    height: 20px;
    margin-top: 15px;
    border-radius: 5px;
  }
  & .react-colorful__saturation-pointer {
    width: 14px;
    height: 14px;
  }
  & .react-colorful__hue-pointer {
    width: 7px;
    border-radius: 3px;
  }
`;

const HashLabel = styled.div`
  margin: auto 0px auto 10px;
  font-size: 13px;
  color: #7d8ec2;
  font-family: ${fonts.mono};
`;

const HexLabel = styled.div`
  margin: auto 10px auto 0px;
  font-size: 12px;
  display: inline-flex;
  font-family: ${fonts.mono};
`;

const HexColorInputBox = styled.div`
  display: inline-flex;
  flex-grow: 1;
  margin-right: 10px;
  width: 140px;
  height: 28px;
  border: 1px solid ${palette.grays.gray3};
  border-radius: 5px;
`;

const HexWrapper = styled.div`
  display: flex;
  flex-grow: 1;
  & input {
    min-width: 0px;
    border: 0px;
    background: ${palette.background.default};
    outline: none;
    font-family: ${fonts.mono};
    font-size: 12px;
    text-transform: uppercase;
    text-align: right;
    padding-right: 10px;
    border-radius: 5px;
  }

  & input:focus {
    border-color: #4298ef;
  }
`;

const Swatch = styled.div<{ $color: string }>`
  display: inline-flex;
  ${({ $color }): string => {
    if ($color.toUpperCase() === '#FFFFFF') {
      return `border: 1px solid ${palette.grays.gray3};`;
    }
    return `border: 1px solid ${$color};`;
  }}
  background-color: ${({ $color }): string => $color};
  width: 28px;
  height: 28px;
  border-radius: 5px;
`;

const ColorPickerInput = styled.div`
  display: flex;
  flex-direction: row;
  align-items: center;
  margin-top: 15px;
`;

export default ColorPicker;
