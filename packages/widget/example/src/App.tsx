import { Workbook } from '@equalto-software/spreadsheet';
import './App.css';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <Workbook type="workbook" initialWorkbook={{ type: 'empty' }} />
      </header>
    </div>
  );
}

export default App;
