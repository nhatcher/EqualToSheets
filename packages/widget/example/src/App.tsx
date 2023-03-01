import * as Workbook from '@equalto-software/spreadsheet';
import './App.css';
import styled from 'styled-components';

function App() {
  const ROWS = 15;
  const COLUMNS = 10;
  return (
    <div className="App">
      <header className="App-header">
        <div style={{ width: COLUMNS * 100 + 30, height: ROWS * 24 + 74, overflow: 'visible' }}>
          <WorkbookRoot lastRow={ROWS + 10} lastColumn={COLUMNS + 10}>
            <Workbook.Toolbar />
            <Workbook.FormulaBar />
            <Workbook.Worksheet />
            <Workbook.Navigation />
          </WorkbookRoot>
        </div>
      </header>
    </div>
  );
}

export default App;

const WorkbookRoot = styled(Workbook.Root)`
  border: 1px solid #c6cae3;
  filter: drop-shadow(0px 2px 2px rgba(33, 36, 58, 0.15));
`;
