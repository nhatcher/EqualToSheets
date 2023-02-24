import styled from 'styled-components/macro';
import { Button, TextField } from '@mui/material';
import { useCallback, useState } from 'react';
import { CheckCircle } from 'lucide-react';

type MailSubmitState =
  | { state: 'waiting' }
  | { state: 'submitting' }
  | { state: 'success' }
  | { state: 'error' };

function useEmailSubmit(): {
  submitState: MailSubmitState;
  setMail: (mail: string) => void;
  submitMail: () => void;
} {
  const [submitState, setSubmitState] = useState<MailSubmitState>({ state: 'waiting' });
  const [mail, setMail] = useState('');
  const submitMail = useCallback(() => {
    setSubmitState({ state: 'submitting' });

    const simpleEmailRegex = /^[^@\s]+@[^@\s]+$/;
    const sanitizedMail = mail.trim();

    if (simpleEmailRegex.test(sanitizedMail)) {
      submitMail(sanitizedMail);
    } else {
      setSubmitState({ state: 'error' });
    }

    async function submitMail(sanitizedMail: string) {
      try {
        const response = await fetch('./signup', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json',
          },
          body: JSON.stringify({ email: sanitizedMail }),
        });

        if (!response.ok) {
          setSubmitState({ state: 'error' });
          return;
        }

        setSubmitState({ state: 'success' });
      } catch {
        setSubmitState({ state: 'error' });
      }
    }
  }, [mail]);

  return { submitState, setMail, submitMail };
}

export function EmailCollectionBar() {
  const { submitState, setMail, submitMail } = useEmailSubmit();
  const formDisabled = submitState.state === 'success' || submitState.state === 'submitting';

  return (
    <Container>
      <SpreadsheetsCopy>
        ✨ <strong>EqualTo Chat</strong> was built with our serverless spreadsheet tech, and
        leverages <strong>GPT-3</strong> learning models.
      </SpreadsheetsCopy>
      {(submitState.state === 'waiting' ||
        submitState.state === 'submitting' ||
        submitState.state === 'error') && (
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
          />
          <SubmitButton type="submit" disabled={formDisabled}>
            Get early access
          </SubmitButton>
          <ConsentText>
            By subscribing I consent to receive communications regarding EqualTo's products and
            services.
          </ConsentText>
        </InlineSignupForm>
      )}
      {submitState.state === 'success' && (
        <SuccessMessageBox>
          <CheckCircle />
          <SuccessText>Thank you for signing up!</SuccessText>
        </SuccessMessageBox>
      )}
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
  font-size: 11px;
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

const EmailInput = styled(TextField).attrs({
  name: 'email',
  placeholder: 'Your email',
})`
  .MuiInputBase-root {
    background: rgba(255, 255, 255, 0.05);
    border-radius: 5px;
    width: 300px;
  }

  input {
    box-sizing: border-box;
    height: 35px;
    padding: 10px;
    color: #f1f2f8;
    font-weight: 400;
    font-size: 13px;
    line-height: 16px;
    &::placeholder {
      color: #8b8fad;
    }
  }
`;

const SubmitButton = styled(Button)`
  box-sizing: border-box;
  height: 35px;
  padding-left: 10px;
  padding-right: 10px;
  margin-left: 5px;

  box-shadow: 0px 2px 2px rgba(0, 0, 0, 0.15);
  background: #70d379;
  border: none;
  border-radius: 5px;

  color: #21243a;
  font-weight: 500;
  font-size: 13px;
  line-height: 16px;
  letter-spacing: 0;
  text-transform: none;

  &:hover {
    background: #59ad60;
    color: #21243a;
  }
`;
