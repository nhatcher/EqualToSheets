import { Dialog, Stack } from '@mui/material';
import styled from 'styled-components/macro';
import { EmailInput } from './emailInput';
import { Button, SubmitButton } from './buttons';
import { useEmailSubmit } from './useEmailSubmit';
import { CheckCircle, X as Close } from 'lucide-react';

export function SignupDialog(properties: { open: boolean; onClose: () => void }) {
  const { open, onClose } = properties;

  const { submitState, setMail, submitMail } = useEmailSubmit();
  const formDisabled = submitState.state === 'success' || submitState.state === 'submitting';

  return (
    <StyledDialog open={open} onClose={onClose}>
      <ButtonContainer>
        <CloseButton onClick={onClose}>
          <Close size={15} />
        </CloseButton>
      </ButtonContainer>
      <Stack pl={6} pr={6} pb={6} spacing={4}>
        <DialogHeading>{'Serverless Spreadsheets'}</DialogHeading>
        <DialogBodyText>
          EqualTo Chat was built with our serverless spreadsheet tech EqualTo Sheets, and OpenAI.
        </DialogBodyText>
        {submitState.state !== 'success' ? (
          <>
            <DialogBodyText>Join the waitlist for early access:</DialogBodyText>
            <DialogForm
              onSubmit={(event) => {
                event.preventDefault();
                submitMail();
              }}
            >
              <EmailInput
                onChange={(event) => setMail(event.target.value)}
                disabled={formDisabled}
                error={submitState.state === 'error'}
              />
              <SubmitButton type="submit" disabled={formDisabled}>
                Join waitlist
              </SubmitButton>
            </DialogForm>
            <DialogConsentText>
              By subscribing I consent to receive communications regarding EqualTo's products and
              services.
            </DialogConsentText>
          </>
        ) : (
          <SuccessMessageBox>
            <CheckCircle />
            <SuccessText>Thank you for joining the waitlist!</SuccessText>
          </SuccessMessageBox>
        )}
      </Stack>
    </StyledDialog>
  );
}

const StyledDialog = styled(Dialog)`
  .MuiPaper-root {
    padding: 0;
    max-width: 360px;
    background: linear-gradient(180deg, #21243a 0%, #292d50 100%, rgba(41, 44, 66, 0.96) 100%);
  }
`;

const DialogHeading = styled.div`
  font-family: var(--monospace-font-family);
  font-weight: 700;
  font-size: 36px;
  line-height: 120%;
  text-align: center;
  color: #b4b7d1;
`;

const DialogBodyText = styled.div`
  font-weight: 400;
  font-size: 20px;
  line-height: 24px;
  text-align: center;
  color: #f1f2f8;
`;

const DialogConsentText = styled.div`
  font-weight: 400;
  font-size: 11px;
  line-height: 13px;
  color: #8b8fad;
  text-align: center;
`;

const SuccessMessageBox = styled.div`
  text-align: center;
  color: #70d379;
  font-size: 16px;
  > svg {
    position: relative;
    top: 6px;
  }
`;

const SuccessText = styled.span`
  margin-left: 10px;
`;

const DialogForm = styled.form`
  display: flex;
  flex-direction: column;
  gap: 10px;
`;

const ButtonContainer = styled.div`
  display: flex;
  justify-content: flex-end;
  padding: 10px;
`;

const CloseButton = styled(Button)`
  width: 40px;
  min-width: 0;
  height: 40px;

  background: rgba(255, 255, 255, 0.1);
  color: #ffffff;
  &:hover {
    background: rgba(255, 255, 255, 0.1);
    color: #ffffff;
  }
`;
