import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import React, { FunctionComponent, ReactNode } from 'react';
import styled from 'styled-components';
import { palette } from 'src/theme';

export type MenuProps = {
  id?: string;
  anchorRef: React.RefObject<HTMLDivElement>;
  onExited?: () => void;
  children: ReactNode;
  className?: string;
};

export const Root: FunctionComponent<{
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
  children: ReactNode;
}> = (properties) => (
  <DropdownMenu.Root open={properties.open} onOpenChange={properties.onOpenChange}>
    {properties.children}
  </DropdownMenu.Root>
);

export const Content: FunctionComponent<{
  children: ReactNode;
  className?: string;
  onExited?: () => void;
}> = (properties) => (
  <DropdownMenu.Portal>
    <StyledContent className={properties.className} onCloseAutoFocus={properties.onExited}>
      {properties.children}
    </StyledContent>
  </DropdownMenu.Portal>
);

export const Item = styled(DropdownMenu.Item)`
  font-size: 14px;
  &:hover,
  &:focus {
    background-color: ${palette.grays.gray1};
  }
  &:active {
    background-color: ${palette.grays.gray2};
  }
  transition: all;
  transition-duration: 200ms;
  padding: 10px;
  cursor: pointer;
  user-select: none;
`;

const StyledContent = styled(DropdownMenu.Content)`
  background: ${palette.background.default};
  box-shadow: 0px 4px 10px rgba(34, 42, 100, 0.1);
  border-radius: 10px;
  padding: 10px;
`;

export const Divider = styled(DropdownMenu.Separator)`
  color: ${palette.grays.gray1};
  background-color: ${palette.grays.gray1};
`;

export const Trigger = styled(DropdownMenu.Trigger)`
  all: unset;
`;
