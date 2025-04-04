import React from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const ExcelExampleExport = () => {
  const data = [
    { name: 'John', age: 30, city: 'New York' },
    { name: 'Jane', age: 25, city: 'San Francisco' },
  ];

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const dataBlob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(dataBlob, 'example.xlsx');
    
    // Add page reload after a short delay to ensure file download completes
    setTimeout(() => {
      window.location.reload();
    }, 1000);
  };

  return (
    <div>
      <h2>Excel Export Example</h2>
      <button onClick={exportToExcel}>Export to Excel</button>
    </div>
  );
};

export default ExcelExampleExport;
