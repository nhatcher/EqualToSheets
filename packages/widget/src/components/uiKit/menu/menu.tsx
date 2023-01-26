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
  &:hover {
    background-color: #ffffff;
    color: ${palette.text.secondary};
  }
  transition: all;
  transition-duration: 200ms;
  padding: '10px 20px 10px 20px';
`;

const StyledContent = styled(DropdownMenu.Content)`
  box-shadow: 0px 11px 15px -7px rgba(54, 62, 125, 0.2), 0px 24px 38px 3px rgba(54, 62, 125, 0.14),
    0px 9px 46px 8px rgba(54, 62, 125, 0.12);
  border-radius: 10px 10px 10px 10px;
  padding: 10px 0px;
  background: #fff;
`;

export const Divider = styled(DropdownMenu.Separator)`
  margin: 0 10px;
  color: ${palette.grays.gray1};
  background-color: ${palette.grays.gray1};
`;

export const Trigger = styled(DropdownMenu.Trigger)`
  all: unset;
`;
