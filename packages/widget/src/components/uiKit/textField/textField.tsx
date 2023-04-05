import React from 'react';
import * as RadixLabel from '@radix-ui/react-label';
import { BodyText } from 'src/components/uiKit/typography';
import styled from 'styled-components';
import { fonts, palette } from 'src/theme';

const TextField: React.FC<{
  name: string;
  className?: string;
  label?: string;
  defaultValue?: string;
  placeholder?: string;
  maxLength?: number;
  autoFocus?: boolean;
  error?: boolean;
  helperText?: string;
  disabled?: boolean;
}> = (properties) => (
  <Container className={properties.className}>
    <Label htmlFor={properties.name}>{properties.label}</Label>
    <Input
      type="text"
      name={properties.name}
      id={properties.name}
      defaultValue={properties.defaultValue}
      maxLength={properties.maxLength}
      // eslint-disable-next-line jsx-a11y/no-autofocus
      autoFocus={properties.autoFocus}
      disabled={properties.disabled}
      placeholder={properties.placeholder}
    />
    {properties.helperText ? <BodyText>{properties.helperText}</BodyText> : null}
  </Container>
);

export default TextField;

const Container = styled.div`
  display: flex;
  flex-direction: column;
`;

const Label = styled(RadixLabel.Root)`
  font-size: 14px;
  margin-bottom: 10px;
`;

const Input = styled.input`
  font-family: ${fonts.regular};
  font-size: 14px;
  box-sizing: border-box;
  &:hover,
  &:focus {
    border: 1px solid ${palette.grays.gray4};
  }
  border: 1px solid ${palette.grays.gray3};
  border-radius: 5px;
  padding: 10px;
`;
