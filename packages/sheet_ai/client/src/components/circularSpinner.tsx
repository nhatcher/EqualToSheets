import styled, { css } from 'styled-components/macro';

// CC0 licensed: https://loading.io/css/
export const CircularSpinner = styled.div<{ $color: string }>`
  display: inline-block;
  width: 80px;
  height: 80px;
  &:after {
    content: ' ';
    display: block;
    width: 64px;
    height: 64px;
    margin: 8px;
    border-radius: 50%;
    ${({ $color }) => css`
      border: 6px solid ${$color};
      border-color: ${$color} transparent ${$color} transparent;
    `}
    animation: lds-dual-ring 1.2s linear infinite;
  }
  @keyframes lds-dual-ring {
    0% {
      transform: rotate(0deg);
    }
    100% {
      transform: rotate(360deg);
    }
  }
`;
