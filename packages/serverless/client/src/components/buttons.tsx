import { Button as MuiButton } from '@mui/material';
import styled from 'styled-components/macro';

export const Button = styled(MuiButton)`
  box-sizing: border-box;
  height: 35px;
  padding-left: 10px;
  padding-right: 10px;

  box-shadow: 0px 2px 2px rgba(0, 0, 0, 0.15);
  border: none;
  border-radius: 5px;

  font-weight: 500;
  font-size: 13px;
  line-height: 16px;
  letter-spacing: 0;
  text-transform: none;

  background: #70d379;
  color: #21243a;
  &:hover {
    background: #59ad60;
    color: #21243a;
  }
`;

export const SubmitButton = styled(Button)`
  background: #70d379;
  color: #21243a;
  &:hover {
    background: #59ad60;
    color: #21243a;
  }
`;
