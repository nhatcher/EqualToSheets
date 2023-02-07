import * as Workbook from '@equalto-software/spreadsheet';
import './App.css';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <Workbook.Root>
          <Workbook.Toolbar />
          <Workbook.FormulaBar />
          <Workbook.Worksheet />
          <Workbook.Navigation />
        </Workbook.Root>
      </header>
    </div>
  );
}

export default App;
