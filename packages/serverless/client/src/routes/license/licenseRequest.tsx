import { FormEventHandler, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import styled from 'styled-components/macro';
import { SubmitButton } from '../../components/buttons';
import { EmailInput } from '../../components/emailInput';
import {
  DualBox,
  ExternalLink,
  HeadingText,
  LeftSide,
  RightSide,
  Subtitle,
  VideoPlaceholder,
} from './common';

export const LicenseRequestPage = () => {
  const navigate = useNavigate();

  const [isDuringSubmit, setDuringSubmit] = useState(false);
  const [email, setEmail] = useState('');
  const [emailError, setEmailError] = useState(false);
  const [requestError, setRequestError] = useState<string | null>(null);

  const onSubmit: FormEventHandler<HTMLFormElement> = (event) => {
    event.preventDefault();

    sendRequest();

    async function sendRequest() {
      const simpleEmailRegex = /^[^@\s]+@[^@\s]+$/;
      const sanitizedMail = email.trim();

      if (!simpleEmailRegex.test(sanitizedMail)) {
        setEmailError(true);
        return;
      }

      setDuringSubmit(true);
      setRequestError(null);
      setEmailError(false);

      const formData = new FormData();
      formData.append('email', sanitizedMail);

      let response;
      try {
        response = await fetch('./send-license-key', {
          method: 'POST',
          headers: {
            Accept: 'application/json',
          },
          body: formData,
        });
      } catch {
        setRequestError('Could not connect to a server. Please try again.');
        setDuringSubmit(false);
        return;
      }

      if (!response.ok) {
        setRequestError(await response.text());
        setDuringSubmit(false);
        return;
      }

      setDuringSubmit(false);
      navigate('/license/sent');
    }
  };

  return (
    <DualBox>
      <LeftSide>
        <HeadingText>
          Get your <em>EqualTo Sheets</em> license key
        </HeadingText>
        <Subtitle>Integrate a high-performance spreadsheet in minutes</Subtitle>
        <VideoPlaceholder>VIDEO</VideoPlaceholder>
      </LeftSide>
      <RightSide>
        <div />
        <Form onSubmit={onSubmit}>
          <EmailInput
            disabled={isDuringSubmit}
            name="email"
            onChange={(event) => setEmail(event.target.value)}
            error={emailError}
            autoFocus
          />
          <SubmitButton disabled={isDuringSubmit} type="submit">
            Get license key
          </SubmitButton>
          {requestError && <ErrorMessage>{requestError}</ErrorMessage>}
        </Form>
        <FormFooterText>
          {'By submitting my details, I agree to the '}
          <ExternalLink href="https://www.equalto.com/tos">Terms of Service</ExternalLink>
          {' and '}
          <ExternalLink href="https://www.equalto.com/privacy-policy">Privacy Policy</ExternalLink>.
        </FormFooterText>
      </RightSide>
    </DualBox>
  );
};

const Form = styled.form`
  display: grid;
  gap: 15px;
  width: 100%;
`;

const FormFooterText = styled.div`
  max-width: 180px;
  text-align: center;
  font-weight: 400;
  font-size: 9px;
  line-height: 11px;
  color: #8b8fad;
`;

const ErrorMessage = styled.div`
  color: #e06276;
  text-align: center;
`;
