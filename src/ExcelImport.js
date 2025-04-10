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

        // Handle merged cells
        const mergedCells = ws['!merges'] || [];
        mergedCells.forEach((merge) => {
          const startRow = merge.s.r;
          const startCol = merge.s.c;
          const endRow = merge.e.r;
          const endCol = merge.e.c;
          const value = rows[startRow]?.[startCol] || '';

          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              if (r === startRow && c === startCol) continue;
              rows[r][c] = value;
            }
          }
        });

        setPreviewData({ headers, rows, file: data });

        const formattedData = rows.map((row) => {
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

        setPreviewData({ headers, rows, file: data });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const openInNewTab = () => {
    const newWindow = window.open();
    newWindow.document.write('<html><head><title>Excel Preview</title></head><body>');
    newWindow.document.write('<style>');
    newWindow.document.write(`
      body { font-family: Arial, sans-serif; padding: 20px; }
      .sheet-tabs { display: flex; gap: 5px; margin-bottom: 10px; }
      .sheet-tab { padding: 8px 16px; border: 1px solid #ddd; cursor: pointer; background: #f5f5f5; }
      .sheet-tab.active { background: #3498db; color: white; }
      table { border-collapse: collapse; width: 100%; margin-top: 10px; }
      th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
      th { background: #f5f5f5; }
    `);
    newWindow.document.write('</style>');

    newWindow.document.write('<div class="sheet-tabs">');
    sheets.forEach((sheet, index) => {
      newWindow.document.write(`
        <button class="sheet-tab ${index === 0 ? 'active' : ''}" onclick="showSheet('${sheet}')">${sheet}</button>
      `);
    });
    newWindow.document.write('</div>');

    sheets.forEach((sheet, index) => {
      const ws = XLSX.utils.sheet_to_json(XLSX.read(previewData.file, { type: 'array' }).Sheets[sheet], { header: 1, defval: '' });
      newWindow.document.write(`<table id="sheet-${sheet}" style="display: ${index === 0 ? 'table' : 'none'};">`);
      newWindow.document.write('<thead><tr>');
      ws[0].forEach((header) => {
        newWindow.document.write(`<th>${header}</th>`);
      });
      newWindow.document.write('</tr></thead><tbody>');
      ws.slice(1).forEach((row) => {
        newWindow.document.write('<tr>');
        row.forEach((cell) => {
          newWindow.document.write(`<td>${cell}</td>`);
        });
        newWindow.document.write('</tr>');
      });
      newWindow.document.write('</tbody></table>');
    });

    newWindow.document.write(`
      <script>
        function showSheet(sheetName) {
          document.querySelectorAll('table').forEach(table => table.style.display = 'none');
          document.querySelectorAll('.sheet-tab').forEach(tab => tab.classList.remove('active'));
          document.getElementById('sheet-' + sheetName).style.display = 'table';
          event.target.classList.add('active');
        }
      </script>
    `);

    newWindow.document.write('</body></html>');
    newWindow.document.close();
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
        <>
          <div className="excel-table-wrapper">
            <table className="excel-table">
              <thead>
                <tr>
                  <th className="row-header corner-header">#</th>
                  {previewData.headers.map((header, index) => (
                    <th key={index} className="column-cell">
                      <div className="column-letter">
                        {String.fromCharCode(65 + index)}
                      </div>
                      <div className="header-content">
                        {header}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {previewData.rows.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    <td className="row-header">{rowIndex + 1}</td>
                    {row.map((cell, cellIndex) => (
                      <td key={cellIndex}>{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button onClick={openInNewTab}>View in New Tab</button>
        </>
      )}
    </div>
  );
};

export default ExcelImport;
