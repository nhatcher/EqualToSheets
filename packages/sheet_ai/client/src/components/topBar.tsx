import styled from 'styled-components/macro';
import { ReactComponent as EqualtoLogo } from './equaltoLogo.svg';

export function TopBar() {
  return (
    <Container>
      <LogoBox>
        <EqualtoLogo height="10px" />
        <Divider />
        <code>Chat</code>
      </LogoBox>
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
