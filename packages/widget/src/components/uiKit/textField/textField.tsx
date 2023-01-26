import React from 'react';
import * as Label from '@radix-ui/react-label';
import { BodyText } from 'src/components/uiKit/typography';

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
  <div className={properties.className}>
    <Label.Root htmlFor={properties.name}>{properties.label}</Label.Root>
    <input
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
  </div>
);

export default TextField;
