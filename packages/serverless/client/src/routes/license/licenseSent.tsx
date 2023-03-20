import styled from 'styled-components/macro';
import { DualBox, LeftSide, RightSide, Subtitle, HeadingText as LeftHeadingText } from './common';
import { ReactComponent as SentSvg } from './sent.svg';
import { VideoEmbed } from './videoEmbed';

export const LicenseSentPage = () => {
  return (
    <DualBox>
      <LeftSide>
        <LeftHeadingText>
          Get your <em>EqualTo Sheets</em> license key
        </LeftHeadingText>
        <Subtitle>Integrate a high-performance spreadsheet in minutes</Subtitle>
        <VideoEmbed />
      </LeftSide>
      <RightSide>
        <div />
        <Container>
          <SentSvg viewBox="0 0 330 230" width={280} />
          <HeadingText>
            Thank you
            <br />
            for signing up!
          </HeadingText>
          <Text>
            You should receive an email shortly, to verify your account and activate your license.
          </Text>
        </Container>
        <div />
      </RightSide>
    </DualBox>
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

const Container = styled.div`
  display: flex;
  align-items: center;
  flex-direction: column;
`;
