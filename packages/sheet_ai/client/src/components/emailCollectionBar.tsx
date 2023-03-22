import styled from 'styled-components/macro';
import { SubmitButton } from './buttons';

export function EmailCollectionBar() {
  return (
    <Container>
      <SpreadsheetsCopy>
        ✨ Built with EqualTo Sheets, our "Spreadsheets as a service" platform for developers, and
        OpenAI.
      </SpreadsheetsCopy>
      <SubmitButton
        sx={{ mr: 4 }}
        type="button"
        onClick={() => {
          window.open('https://sheets.equalto.com/', '_blank');
        }}
      >
        Access EqualTo Sheets now
      </SubmitButton>
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
