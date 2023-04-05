import * as RadixDialog from '@radix-ui/react-dialog';
import React, { ReactNode } from 'react';
import Stack from 'src/components/uiKit/stack';
import { BodyText, DialogTitle } from 'src/components/uiKit/typography';
import { XIcon } from 'lucide-react';
import styled from 'styled-components';
import * as CSS from 'csstype';
import StylelessButton from 'src/components/uiKit/button/styleless';
import { fonts, palette } from 'src/theme';

export const DialogHeader = styled.div`
  display: flex;
  flex-shrink: 0;
  justify-content: space-between;
  padding: 0 30px 30px 30px;
`;

const DialogTitleContainer = styled.div`
  display: inline-block;
`;

const DialogContentContainer = styled(RadixDialog.Content)<{
  $width?: CSS.Property.Width;
  $height?: CSS.Property.Height;
  $maxHeight?: CSS.Property.MaxHeight;
}>`
  font-family: ${fonts.regular};
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background: #fff;
  max-width: ${({ $width }) => $width ?? '450px'};
  width: ${({ $width }) => $width ?? '450px'};
  ${({ $height }) => ($height ? `height: ${$height};` : '')};
  ${({ $maxHeight }) => ($maxHeight ? `max-height: ${$maxHeight};` : '')};
  box-shadow: 0px 11px 15px -7px rgba(54, 62, 125, 0.2), 0px 24px 38px 3px rgba(54, 62, 125, 0.14),
    0px 9px 46px 8px rgba(54, 62, 125, 0.12);
  border-radius: 10px;
  padding-top: 30px;
`;

export const DialogContent = styled.div<{ $overflowContent?: boolean }>`
  ${({ $overflowContent }) => ($overflowContent ? 'overflow: auto;' : '')};
  display: flex;
  flex-direction: column;
  flex: 1;
  padding: 0 30px 30px 30px;
`;

const IconSlot = styled.div`
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  max-width: 50px;
  max-height: 50px;
  overflow: hidden;
`;

type DialogProps = {
  className?: string;
  open: boolean;
  /** Callback fired when the component requests to be closed. */
  onClose: () => void;
  /** Callback fired when the dialog has exited. - Animation has finished. */
  onExited?: () => void;
  title: string | null | JSX.Element;
  description?: string;
  closeTitle?: string | null;
  keepMounted?: boolean;
  icon?: ReactNode;
  children?: ReactNode;
  $width?: CSS.Property.Width;
  $height?: CSS.Property.Height;
  $maxHeight?: CSS.Property.MaxHeight;
  $overflowContent?: boolean;
};

const Dialog: React.FC<DialogProps> = (properties) => (
  <RadixDialog.Root open={properties.open} onOpenChange={properties.onClose}>
    <RadixDialog.Portal>
      <RadixDialog.Trigger />
      <RadixDialog.Overlay />
      <DialogContentContainer
        className={properties.className}
        onCloseAutoFocus={properties.onExited}
        $width={properties.$width}
        $height={properties.$height}
        $maxHeight={properties.$maxHeight}
      >
        {properties.title !== null && (
          <DialogHeader>
            <Stack $direction="row" $align="center">
              {properties.icon && <IconSlot>{properties.icon}</IconSlot>}
              <DialogTitleContainer>
                <DialogTitle>{properties.title}</DialogTitle>
                {properties.description && (
                  <DescriptionText>{properties.description}</DescriptionText>
                )}
              </DialogTitleContainer>
            </Stack>
            <CloseButton onClick={properties.onClose} aria-label="Close">
              <XIcon size={14} />
            </CloseButton>
          </DialogHeader>
        )}
        <DialogContent $overflowContent={properties.$overflowContent}>
          {properties.children}
        </DialogContent>
      </DialogContentContainer>
    </RadixDialog.Portal>
  </RadixDialog.Root>
);

const DescriptionText = styled(BodyText)`
  color: ${palette.grays.gray4};
`;
const CloseButton = styled(StylelessButton)`
  display: inline-flex;
  justify-content: center;
  align-items: center;
  font-size: 15px;
  color: #b4b7d1;
  width: 20px;
  height: 20px;
  transition: color 0.2s;

  &:focus {
    color: #8b8fad;
  }

  &:not(:disabled) {
    &:hover {
      color: #8b8fad;
    }
    &:active {
      color: #5f6989;
    }
  }

  &:disabled {
    color: ${palette.grays.gray3};
  }
`;

export default Dialog;
