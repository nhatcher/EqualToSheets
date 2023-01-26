import { lighten } from 'polished';
import styled from 'styled-components';
import { palette } from 'src/theme';
import StylelessButton from './styleless';

const LightButton = styled(StylelessButton)<{
  $backgroundHighlight?: string;
  $sizePx?: number;
}>`
  width: ${({ $sizePx }): number => $sizePx ?? 25}px;
  height: ${({ $sizePx }): number => $sizePx ?? 25}px;
  font-size: ${({ $sizePx }): number => Math.floor((2 * ($sizePx ?? 25)) / 3)}px;

  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;

  border-radius: 5px;
  color: ${palette.grays.gray3};
  transition-duration: 0.3s;
  transition-property: background-color, color;

  &:focus {
    background: ${({ $backgroundHighlight }): string =>
      lighten(0.01, $backgroundHighlight ?? palette.grays.gray2)};
  }

  &:not(:disabled) {
    &:hover {
      background: ${({ $backgroundHighlight }): string =>
        $backgroundHighlight ?? palette.grays.gray2};
    }
    &:active {
      background: ${({ $backgroundHighlight }): string =>
        $backgroundHighlight ?? palette.grays.gray2};
      color: ${palette.grays.gray4};
    }
  }

  &:disabled {
    color: ${palette.grays.gray3};
    pointer-events: none;
  }
`;

export default LightButton;
