import React from 'react';
import ExcelExampleExport from './ExcelExampleExport';
import ExcelImport from './ExcelImport';

function App() {
  return (
    <div className="App">
      <h1>Excel Import/Export Demo</h1>
      <div style={{ margin: '20px' }}>
        <ExcelExampleExport />
      </div>
      <div style={{ margin: '20px' }}>
        <ExcelImport />
      </div>
    </div>
  );
}

export default App;
