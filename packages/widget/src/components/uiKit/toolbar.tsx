import { palette } from 'src/theme';
import * as RadixToolbar from '@radix-ui/react-toolbar';
import styled from 'styled-components';

export const Separator = styled(RadixToolbar.Separator)`
  display: inline-flex;
  height: 10px;
  width: 1px;
  border-left: 1px solid ${palette.grays.gray2};
  margin-left: 5px;
  margin-right: 5px;
`;

export const Root = styled(RadixToolbar.Root)`
  display: flex;
  flex-shrink: 0;
  align-items: center;
  background: ${palette.background.default};
  height: 40px;
  border-bottom: 1px solid ${palette.grays.gray2};
`;

export const Button = styled(RadixToolbar.Button)<{
  $underlinedColor?: string;
  $selected?: boolean;
}>`
  all: unset;

  box-sizing: border-box;
  cursor: ${({ disabled }): string => (!disabled ? 'pointer' : 'default')};
  width: 23px;
  height: 23px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border-radius: 2px;
  margin-right: 5px;
  transition: all 0.2s;

  font-size: 12px;
  font-weight: 600;

  ${({ disabled, $underlinedColor, $selected }): string => {
    if (disabled) {
      return `
      color: ${palette.grays.gray3};
    `;
    }
    return `
      border-top: ${$underlinedColor ? '3px solid #FFF' : 'none'};
      border-bottom: ${$underlinedColor ? `3px solid ${$underlinedColor}` : 'none'};
      color: ${palette.text.primary};
      background-color: ${$selected ? palette.grays.gray2 : '#FFF'};
      &:hover {
        background-color: ${palette.grays.gray1};
        border-top-color: ${palette.grays.gray1};
      }
    `;
  }}
`;
