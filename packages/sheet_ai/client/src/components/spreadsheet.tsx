import styled from 'styled-components/macro';

export const Spreadsheet = () => {
  return <SpreadsheetPlaceholder />;
};

const SpreadsheetPlaceholder = styled.div`
  height: 250px;
  background: repeating-linear-gradient(-45deg, #ffffff, #ffffff 10px, #d2d2f2 10px, #d2d2f2 20px);
`;
