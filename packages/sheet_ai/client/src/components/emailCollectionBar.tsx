import { CheckCircle } from 'lucide-react';
import { useState } from 'react';
import styled from 'styled-components/macro';
import { Button, SubmitButton } from './buttons';
import { EmailInput } from './emailInput';
import { SignupDialog } from './signupDialog';
import { useEmailSubmit } from './useEmailSubmit';

export function EmailCollectionBar() {
  const { submitState, setMail, submitMail } = useEmailSubmit();
  const formDisabled = submitState.state === 'success' || submitState.state === 'submitting';

  const [signupDialogOpen, setSignupDialogOpen] = useState(false);

  return (
    <Container>
      <SpreadsheetsCopy>
        ✨ Built with our serverless spreadsheet tech EqualTo Sheets, and OpenAI.
      </SpreadsheetsCopy>
      {(submitState.state === 'waiting' ||
        submitState.state === 'submitting' ||
        submitState.state === 'error') && (
        <>
          <InlineSignupForm
            onSubmit={(event) => {
              event.preventDefault();
              submitMail();
            }}
          >
            <EmailInput
              onChange={(event) => setMail(event.target.value)}
              disabled={formDisabled}
              error={submitState.state === 'error'}
              $width={300}
            />
            <SubmitButton sx={{ ml: 1 }} type="submit" disabled={formDisabled}>
              Join EqualTo Sheets waitlist
            </SubmitButton>
            <ConsentText>
              By subscribing I consent to receive communications regarding EqualTo's products and
              services.
            </ConsentText>
          </InlineSignupForm>
          <MobileSignupContainer>
            <MobileSignupButton type="button" onClick={() => setSignupDialogOpen(true)}>
              Join waitlist
            </MobileSignupButton>
          </MobileSignupContainer>
        </>
      )}
      {submitState.state === 'success' && (
        <SuccessMessageBox>
          <CheckCircle />
          <SuccessText>Thank you for joining!</SuccessText>
        </SuccessMessageBox>
      )}
      <SignupDialog open={signupDialogOpen} onClose={() => setSignupDialogOpen(false)} />
    </Container>
  );
}

const Container = styled.div`
  display: flex;
  justify-content: space-between;
  background: linear-gradient(180deg, #21243a 0%, #292d50 100%, rgba(41, 44, 66, 0.96) 100%);
  color: #f1f2f8;
  min-height: 75px;
  align-items: center;
`;

const SpreadsheetsCopy = styled.div`
  flex: 1;
  padding: 20px;

  p {
    margin: 0;
  }

  > p + p {
    margin-top: 5px;
  }

  font-family: var(--font-family);
  color: #f1f2f8;
  font-weight: 400;
  font-size: 13px;
  line-height: 13px;
`;

const InlineSignupForm = styled.form`
  align-self: center;
  padding: 0 20px;
  position: relative;
  @media (max-width: 800px) {
    display: none;
  }
`;

const ConsentText = styled.div`
  margin-top: 2px;
  font-weight: 400;
  font-size: 7px;
  line-height: 8px;
  color: #8b8fad;
  position: absolute;
  top: 100%;
`;

const SuccessMessageBox = styled.div`
  display: flex;
  align-items: center;
  padding-right: 20px;
  color: #70d379;
  @media (max-width: 800px) {
    display: none;
  }
`;

const SuccessText = styled.div`
  margin-left: 10px;
`;

const MobileSignupContainer = styled.div`
  padding: 0 20px;
  @media (min-width: 801px) {
    display: none;
  }
`;

const MobileSignupButton = styled(Button)`
  height: 40px;
  font-weight: 400;
  font-size: 14px;
  line-height: 17px;
  color: #ffffff;
  background: rgba(255, 255, 255, 0.2);
  &:hover {
    color: #ffffff;
    background: rgba(255, 255, 255, 0.3);
  }
`;
