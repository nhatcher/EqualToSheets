import { transparentize } from 'polished';
import React, { ReactNode, ForwardRefExoticComponent, forwardRef, Ref } from 'react';
import { ButtonText } from 'src/components/uiKit/typography';
import styled from 'styled-components';
import { palette } from 'src/theme';
import StylelessButton from './styleless';

export type ButtonProps = {
  id?: string;
  className?: string;
  children?: ReactNode;
  variant?:
    | 'primary'
    | 'primary-red'
    | 'primary-blue'
    | 'primary-orange'
    | 'primary-alt-green'
    | 'primary-alt-red'
    | 'disabled'
    | 'secondary';
  onClick?: React.MouseEventHandler<HTMLButtonElement>;
  width?: string;
  height?: string;
  disabled?: boolean;
  type?: 'button' | 'submit' | 'reset';
  StartIcon?: ReactNode;
  EndIcon?: ReactNode;
  'aria-label'?: string;
  'data-testid'?: string;
  ref?: Ref<HTMLButtonElement>;
};

const ButtonBase: ForwardRefExoticComponent<ButtonProps> = forwardRef<
  HTMLButtonElement,
  ButtonProps
>((properties, reference) => {
  const { className, children, type, StartIcon, EndIcon, onClick, disabled, id } = properties;
  return (
    <StylelessButton
      id={id}
      ref={reference}
      className={className}
      onClick={onClick}
      type={type}
      disabled={disabled}
      aria-label={properties['aria-label']}
      data-testid={properties['data-testid']}
    >
      {StartIcon}
      <ButtonText>{children}</ButtonText>
      {EndIcon}
    </StylelessButton>
  );
});

const ButtonPrimary = styled(ButtonBase)<ButtonProps>`
  background-color: ${palette.primary.dark};
  color: ${palette.background.default};
  &:focus,
  &:hover {
    background-color: #62c16a;
  }
  &:disabled {
    background-color: ${transparentize(0.25, palette.text.disabled)};
    color: ${palette.text.disabled};
  }
`;

const ButtonPrimaryRed = styled(ButtonBase)<ButtonProps>`
  background-color: ${palette.error.main};
  color: ${palette.background.default};
  &:focus,
  &:hover {
    background-color: #d65065;
  }
  &:disabled {
    background-color: ${transparentize(0.25, palette.text.disabled)};
    color: ${palette.text.disabled};
  }
`;

const ButtonPrimaryOrange = styled(ButtonBase)<ButtonProps>`
  background-color: ${palette.warning.main};
  color: ${palette.background.default};
  &:focus,
  &:hover {
    background-color: #f2af2e;
  }
  &:disabled {
    background-color: ${transparentize(0.25, palette.text.disabled)};
    color: ${palette.text.disabled};
  }
`;

const ButtonPrimaryBlue = styled(ButtonBase)<ButtonProps>`
  background-color: ${palette.text.link};
  color: ${palette.background.default};
  &:focus,
  &:hover {
    background-color: #2d55e3;
  }
  &:disabled {
    background-color: ${transparentize(0.25, palette.text.disabled)};
    color: ${palette.text.disabled};
  }
`;

const ButtonPrimaryAltGreen = styled(ButtonBase)<ButtonProps>`
  background-color: ${palette.primary.light};
  color: ${palette.primary.darker};
  border: 1px solid ${palette.primary.dark};
  &:focus,
  &:hover {
    background-color: #fafdfb;
  }
  &:disabled {
    background-color: #ffffff;
    border-color: ${transparentize(0.25, palette.text.disabled)};
    color: ${palette.text.disabled};
  }
`;

const ButtonPrimaryAltRed = styled(ButtonBase)<ButtonProps>`
  background-color: #fceff1;
  color: #f05858;
  border: 1px solid #f05858;
  &:focus,
  &:hover {
    background-color: #fffbfb;
  }
  &:disabled {
    background-color: #ffffff;
    border-color: ${transparentize(0.25, palette.text.disabled)};
    color: ${palette.text.disabled};
  }
`;

const ButtonSecondary = styled(ButtonBase)<ButtonProps>`
  background-color: #fff;
  color: ${palette.text.primary};
  border: 1px solid ${palette.grays.gray4};
  &:focus,
  &:hover {
    background-color: ${palette.grays.gray1};
    border: 1px solid ${palette.grays.gray4};
  }
  &:disabled {
    border: 1px solid ${palette.grays.gray3};
    color: ${palette.grays.gray3};
  }
`;

const ButtonDisabled = styled(ButtonBase)<ButtonProps>`
  background-color: ${transparentize(0.25, palette.text.disabled)};
  color: ${palette.text.disabled};
`;

const variantToButton = (variant: string | undefined): ForwardRefExoticComponent<ButtonProps> => {
  switch (variant) {
    case 'primary':
      return ButtonPrimary;
    case 'primary-red':
      return ButtonPrimaryRed;
    case 'primary-blue':
      return ButtonPrimaryBlue;
    case 'primary-orange':
      return ButtonPrimaryOrange;
    case 'primary-alt-green':
      return ButtonPrimaryAltGreen;
    case 'primary-alt-red':
      return ButtonPrimaryAltRed;
    case 'secondary':
      return ButtonSecondary;
    case 'disabled':
      return ButtonDisabled;
    default:
      return ButtonPrimary;
  }
};

const Button: ForwardRefExoticComponent<ButtonProps> = forwardRef<HTMLButtonElement, ButtonProps>(
  (properties, reference) => {
    const { id, className, children, variant, type, StartIcon, EndIcon, onClick, disabled } =
      properties;
    const ButtonVariant = variantToButton(variant);
    const isDisabled = disabled || variant === 'disabled';
    return (
      <ButtonVariant
        id={id}
        ref={reference}
        className={className}
        StartIcon={StartIcon}
        EndIcon={EndIcon}
        onClick={onClick}
        type={type}
        disabled={isDisabled}
        aria-label={properties['aria-label']}
        data-testid={properties['data-testid']}
      >
        {children}
      </ButtonVariant>
    );
  },
);

const StyledButton = styled(Button)`
  border-radius: 5px;
  box-sizing: border-box;
  width: ${(properties): string => properties.width || '200px'};
  height: ${(properties): string => properties.height || '35px'};
  transition: background-color 150ms ease-out;
`;

export default StyledButton;
