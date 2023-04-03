import * as Workbook from '@equalto-software/spreadsheet';
import './App.css';
import styled from 'styled-components';
import '@fontsource/inter/400.css';
import '@fontsource/inter/500.css';
import '@fontsource/inter/600.css';
import '@fontsource/fira-mono/400.css';
import '@fontsource/fira-mono/500.css';
import '@fontsource/fira-mono/700.css';

function App() {
  const ROWS = 15;
  const COLUMNS = 10;
  return (
    <div className="App">
      <header className="App-header">
        <div style={{ width: COLUMNS * 100 + 30, height: ROWS * 24 + 74, overflow: 'visible' }}>
          <WorkbookRoot>
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
  border-radius: 5px;
  filter: drop-shadow(0px 2px 2px rgba(33, 36, 58, 0.15));
`;
