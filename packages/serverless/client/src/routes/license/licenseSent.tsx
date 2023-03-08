import styled from 'styled-components/macro';
import { Box } from './common';
import { ReactComponent as SentSvg } from './sent.svg';

export const LicenseSentPage = () => {
  return (
    <Box $maxWidth={360}>
      <SentSvg />
      <HeadingText>
        Thank you
        <br />
        for signing up!
      </HeadingText>
      <Text>
        You should receive an email shortly, to verify your account and activate your license.
      </Text>
    </Box>
  );
};

const HeadingText = styled.div`
  font-weight: 600;
  font-size: 24px;
  line-height: 29px;
  text-align: center;
  color: #ffffff;
`;

const Text = styled.p`
  font-weight: 400;
  font-size: 14px;
  line-height: 17px;
  text-align: center;
  color: #b4b7d1;
`;
