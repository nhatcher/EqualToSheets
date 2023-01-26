import styled, { css, FlattenSimpleInterpolation } from 'styled-components';
import * as CSS from 'csstype';

// Component inspired by: https://chakra-ui.com/docs/layout/stack

export type StackProps = {
  $spacing?: string;
  $direction?: 'row' | 'column';
  $align?: CSS.Property.AlignItems;
  $justify?: CSS.Property.JustifyContent;
};

const Stack = styled.div<StackProps>`
  display: flex;
  align-items: ${(properties): string => properties.$align || 'stretch'};
  justify-content: ${(properties): string => properties.$justify || 'start'};
  ${(properties): FlattenSimpleInterpolation => {
    switch (properties.$direction) {
      case 'row':
        return css`
          flex-direction: row;

          > *:not(style) ~ *:not(style) {
            margin-left: ${properties.$spacing};
          }
        `;
      case 'column':
      default:
        return css`
          flex-direction: column;

          > *:not(style) ~ *:not(style) {
            margin-top: ${properties.$spacing};
          }
        `;
    }
  }}
`;

Stack.defaultProps = {
  $spacing: '20px',
  $direction: 'row',
};

export const Spacer = styled.div`
  flex: 1 1 0;
  place-self: stretch;
`;

export default Stack;
