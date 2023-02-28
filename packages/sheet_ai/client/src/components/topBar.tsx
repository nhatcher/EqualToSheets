import styled from 'styled-components/macro';
import { ReactComponent as EqualtoLogo } from './equaltoLogo.svg';

export function TopBar() {
  return (
    <Container>
      <LogoBox>
        <UnstyledLink href="https://www.equalto.com/" target="_blank">
          <EqualtoLogo height="10px" />
        </UnstyledLink>
        <Divider />
        <code>Chat</code>
      </LogoBox>
      <Link href="https://www.equalto.com/" target="_blank">
        About EqualTo
      </Link>
    </Container>
  );
}

const Container = styled.div`
  display: flex;
  flex-direction: row;
  align-items: center;
  justify-content: space-between;
  height: 50px;
  padding: 0 20px;
  border-bottom: 1px solid #dee0ef;
`;

const LogoBox = styled.div`
  display: flex;
  flex-direction: row;
  align-items: center;
  justify-content: space-between;
`;

const Divider = styled.div`
  border-right: 1px solid #f1f2f8;
  height: 12px;
  margin: 0 10px;
`;

const UnstyledLink = styled.a`
  :link,
  :visited,
  :hover,
  :active {
    text-decoration: none;
  }
`;

const Link = styled.a`
  font-weight: 400;
  font-size: 13px;
  line-height: 16px;
  :link,
  :visited,
  :hover,
  :active {
    color: #000000;
    text-decoration: none;
  }
`;
