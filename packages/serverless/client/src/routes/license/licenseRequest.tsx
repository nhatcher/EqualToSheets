import { FormEventHandler, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import styled from 'styled-components/macro';
import { SubmitButton } from '../../components/buttons';
import { EmailInput } from '../../components/emailInput';
import { ExternalLink } from './common';

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
        <Subtitle>Start building in minutes</Subtitle>
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

const DualBox = styled.div`
  display: grid;
  grid-template-columns: 6fr 4fr;

  align-self: center;
  width: 100%;
  max-width: 960px;

  border: 1px solid #46495e;
  filter: drop-shadow(0px 2px 4px rgba(0, 0, 0, 0.1));
  border-radius: 16px;
`;

const LeftSide = styled.div`
  padding: 50px;
  background: rgba(255, 255, 255, 0.03);
  box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
`;

const RightSide = styled.div`
  padding: 20px 50px;
  filter: drop-shadow(0px 2px 4px rgba(0, 0, 0, 0.1));
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  align-items: center;
`;

const HeadingText = styled.h1`
  font-style: normal;
  font-weight: 600;
  font-size: 28px;
  line-height: 34px;
  color: #ffffff;
  margin: 0;
  em {
    font-style: normal;
    color: #72ed79;
  }
`;

const Subtitle = styled.p`
  margin: 10px 0 0 0;
  font-weight: 400;
  font-size: 16px;
  line-height: 19px;
  display: flex;
  color: #b4b7d1;
`;

const VideoPlaceholder = styled.div`
  margin-top: 30px;
  background: #f2f2f2;
  border-radius: 10px;
  width: 100%;
  height: 250px;
  text-align: center;
  padding: 40px;
`;

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
