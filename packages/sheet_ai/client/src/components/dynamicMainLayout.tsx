import { ReactNode } from 'react';
import styled from 'styled-components/macro';

export const DynamicMainLayout = (properties: {
  mainContent: ReactNode;
  sideContent: ReactNode;
}) => {
  const { mainContent, sideContent } = properties;
  return (
    <Layout>
      <MainContent>{mainContent}</MainContent>
      <SideContent>{sideContent}</SideContent>
    </Layout>
  );
};

const Layout = styled.div`
  display: flex;
  max-width: 830px;
  width: 100%;
  justify-self: center;
  min-height: 0;
`;

const MainContent = styled.div`
  flex-shrink: 1;
  flex-grow: 1;
`;

const SideContent = styled.div`
  flex-basis: 260px;
  flex-grow: 0;
  flex-shrink: 0;
  border-left: 1px solid #dee0ef;
  margin-left: 30px;
  padding: 30px;

  @media (max-width: 800px) {
    display: none;
  }
`;
