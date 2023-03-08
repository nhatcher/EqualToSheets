import { Outlet } from 'react-router-dom';
import styled from 'styled-components/macro';
import { ReactComponent as EqualToLogo } from '../../logos/equaltoLogo.svg';
import { ExternalLink } from './common';

export const LicenseLayout = () => {
  return (
    <Layout>
      <Heading>
        {/* eslint-disable-next-line react/jsx-no-target-blank */}
        <a href="https://www.equalto.com/" target="_blank">
          <EqualToLogo />
        </a>
      </Heading>
      <Content>
        <Outlet />
      </Content>
      <Footer>
        &copy; Copyright {new Date().getFullYear()} EqualTo GmbH
        {/* eslint-disable-next-line react/jsx-no-target-blank */}
        <StyledExtLink href="https://www.equalto.com/" target="_blank">
          equalto.com
        </StyledExtLink>
        {/* eslint-disable-next-line react/jsx-no-target-blank */}
        <StyledExtLink href="mailto:support@equalto.com">Support</StyledExtLink>
      </Footer>
    </Layout>
  );
};

const Layout = styled.div`
  display: flex;
  flex-direction: column;
  min-height: 100%;
  background: lightblue;

  background: radial-gradient(
      70.43% 179.85% at 21.99% 27.53%,
      rgba(88, 121, 240, 0.07) 0%,
      rgba(88, 121, 240, 0) 100%
    ),
    linear-gradient(180deg, #292c42 0%, #1f2236 100%);
  background-blend-mode: normal, normal;
  box-shadow: 0px 2px 2px rgba(52, 56, 85, 0.15);
`;

const Heading = styled.div`
  display: flex;
  flex-direction: row;
  justify-content: center;
  padding: 30px;
`;

const Content = styled.div`
  flex-grow: 1;
  flex-shrink: 0;
  min-height: min-content;

  display: flex;
  flex-direction: column;
  justify-content: center;
`;

const Footer = styled.div`
  display: flex;
  flex-direction: row;
  justify-content: center;
  gap: 20px;
  padding: 30px;

  font-weight: 400;
  font-size: 9px;
  line-height: 11px;
  color: #8b8fad;
`;

const StyledExtLink = styled(ExternalLink)`
  font-weight: 400;
  font-size: 9px;
  line-height: 11px;
`;
