import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

const ExcelImport = ({ uploadHandler }) => {
  const [errorMessage, setErrorMessage] = useState('');
  const [previewData, setPreviewData] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const handleFile = useCallback((file) => {
    if (!file) {
      setErrorMessage('No file uploaded!');
      return;
    }

    if (!(file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.type === 'application/vnd.ms-excel')) {
      setErrorMessage('Unknown file format. Only Excel files are supported.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      setSheets(workbook.SheetNames);
      setSelectedSheet(workbook.SheetNames[0]);
      
      const ws = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      
      if (jsonData.length) {
        const headers = jsonData[0];
        const rows = jsonData.slice(1);
        
        // Ensure all rows have same length as headers
        const normalizedRows = rows.map(row => {
          const normalizedRow = [...row];
          while (normalizedRow.length < headers.length) {
            normalizedRow.push('');
          }
          return normalizedRow;
        });
        
        setPreviewData({ headers, rows: normalizedRows });
        
        const formattedData = normalizedRows.map(row => {
          let obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index] || '';
          });
          return obj;
        });
        uploadHandler(formattedData);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    
    const file = e.dataTransfer.files[0];
    handleFile(file);
  }, [handleFile]);

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const fileHandler = (event) => {
    const file = event.target.files[0];
    handleFile(file);
  };

  const handleSheetChange = (event) => {
    const sheetName = event.target.value;
    setSelectedSheet(sheetName);
    
    const file = document.querySelector('input[type="file"]').files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const ws = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      
      if (jsonData.length) {
        const headers = jsonData[0];
        const rows = jsonData.slice(1);
        
        // Ensure all rows have same length as headers
        const normalizedRows = rows.map(row => {
          const normalizedRow = [...row];
          while (normalizedRow.length < headers.length) {
            normalizedRow.push('');
          }
          return normalizedRow;
        });
        
        setPreviewData({ headers, rows: normalizedRows });
        
        const formattedData = normalizedRows.map(row => {
          let obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index] || '';
          });
          return obj;
        });
        uploadHandler(formattedData);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="excel-import-container">
      <h2>Import Excel File</h2>
      <div 
        className={`file-upload-area ${isDragging ? 'dragging' : ''}`}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
      >
        <div className="file-upload">
          <input
            type="file"
            onChange={fileHandler}
            accept=".xlsx, .xls"
          />
        </div>
        <div className="drag-drop-text">
          Drag and drop Excel file here or click to browse
        </div>
      </div>

      {sheets.length > 0 && (
        <div className="sheet-tabs">
          {sheets.map((sheet, index) => (
            <button
              key={index}
              className={`sheet-tab ${selectedSheet === sheet ? 'active' : ''}`}
              onClick={() => handleSheetChange({ target: { value: sheet } })}
            >
              {sheet}
            </button>
          ))}
        </div>
      )}
      
      {errorMessage && <p style={{ color: 'red' }}>{errorMessage}</p>}
      
      {previewData && (
        <div className="excel-table-wrapper">
          <table className="excel-table">
            <thead>
              <tr>
                {previewData.headers.map((header, index) => (
                  <th key={index}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewData.rows.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, cellIndex) => (
                    <td key={cellIndex}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default ExcelImport;
