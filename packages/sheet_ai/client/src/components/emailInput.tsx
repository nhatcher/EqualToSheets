import styled, { css } from 'styled-components/macro';
import { TextField } from '@mui/material';

export const EmailInput = styled(TextField).attrs({
  name: 'email',
  placeholder: 'Your email',
})<{ $width?: number }>`
  .MuiInputBase-root {
    background: rgba(255, 255, 255, 0.05);
    border-radius: 5px;
    ${({ $width }) =>
      $width
        ? css`
            width: ${$width}px;
          `
        : ''}
  }

  input {
    box-sizing: border-box;
    height: 35px;
    padding: 10px;
    color: #f1f2f8;
    font-weight: 400;
    font-size: 13px;
    line-height: 16px;
    &::placeholder {
      color: #8b8fad;
    }
  }
`;
