import React, { FunctionComponent, useState } from 'react';
import Dialog from 'src/components/uiKit/dialog';
import styled from 'styled-components';
import Button from 'src/components/uiKit/button';
import TextField from 'src/components/uiKit/textField';
import Constants from 'src/constants';

export type PromptDialogSubmitResult =
  | { success: true }
  | { success: false; errorMessage: string; userFriendlyErrorMessage?: string };
export interface PromptDialogProps {
  open: boolean;
  onClose: () => void;
  onSubmit: (newValue: string) => Promise<PromptDialogSubmitResult>;
  title: string;
  label: string;
  placeholder?: string;
  defaultValue?: string;
  submit?: string;
  submitting?: string;
  cancel?: string;
  required?: boolean;
  maxLength?: number;
  requiredErrorMessage?: string;
}

interface ErrorMessage {
  message: string;
  userFriendlyMessage: string;
}

const PromptDialog: FunctionComponent<PromptDialogProps> = ({
  defaultValue = '',
  ...properties
}) => {
  const [duringSubmit, setDuringSubmit] = useState(false);
  const [errorMessage, setErrorMessage] = useState<null | ErrorMessage>(null);

  const { open, onClose, label, title, onSubmit, required } = properties;

  const placeholder = properties.placeholder ?? 'common.prompt_dialog.placeholder';
  const submit = properties.submit ?? 'common.prompt_dialog.submit';
  const submitting = properties.submitting ?? 'common.prompt_dialog.submitting';
  const cancel = properties.cancel ?? 'common.prompt_dialog.cancel';
  const requiredErrorMessage = properties.requiredErrorMessage ?? 'common.prompt_dialog.required';
  const maxLength = properties.maxLength ?? Constants.CHAR_MAX_LEN;

  return (
    <StyledDialog
      title={title}
      open={open}
      onClose={onClose}
      closeTitle={null}
      onExited={() => {
        setErrorMessage(null);
      }}
    >
      <form
        onSubmit={async (event) => {
          const formData = new FormData(event.target as HTMLFormElement);
          const value = formData.get('value') as string;

          if (required && value.trim().length === 0) {
            setErrorMessage({
              message: requiredErrorMessage,
              userFriendlyMessage: requiredErrorMessage,
            });
            return;
          }

          setDuringSubmit(true);
          const result = await onSubmit(formData.get('value') as string);
          setDuringSubmit(false);
          if (result.success === true) {
            setErrorMessage(null);
            onClose();
          } else {
            setErrorMessage({
              message: result.errorMessage,
              userFriendlyMessage: result.userFriendlyErrorMessage ?? 'error.unexpected',
            });
          }
        }}
      >
        <TextField
          name="value"
          defaultValue={defaultValue}
          label={label}
          maxLength={maxLength}
          disabled={duringSubmit}
          error={errorMessage !== null}
          helperText={(errorMessage && errorMessage.userFriendlyMessage) ?? ''}
          placeholder={placeholder}
          autoFocus
        />
        {errorMessage && errorMessage.userFriendlyMessage}
        <ButtonsContainer>
          <StyledButton type="button" variant="secondary" onClick={onClose} disabled={duringSubmit}>
            {cancel}
          </StyledButton>
          <StyledButton type="submit" variant="primary" disabled={duringSubmit}>
            {!duringSubmit ? submit : submitting}
          </StyledButton>
        </ButtonsContainer>
      </form>
    </StyledDialog>
  );
};

const StyledDialog = styled(Dialog)`
  min-width: 600px;
`;

const ButtonsContainer = styled.div`
  display: flex;
  justify-content: flex-end;
  grid-column-gap: 10px;
  margin-top: 30px;
`;

const StyledButton = styled(Button)`
  padding: 0px 15px;
  width: auto;
`;

export default PromptDialog;
