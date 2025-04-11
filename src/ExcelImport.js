import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Dialog, DialogContent, DialogTitle, IconButton, AppBar, Toolbar, Typography } from '@mui/material';
import CloseIcon from '@mui/icons-material/Close';
import PreviewIcon from '@mui/icons-material/Visibility';

const ExcelImport = ({ uploadHandler }) => {
  const [errorMessage, setErrorMessage] = useState('');
  const [previewData, setPreviewData] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [openPreview, setOpenPreview] = useState(false);

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

  const handlePreviewOpen = () => setOpenPreview(true);
  const handlePreviewClose = () => {
    setOpenPreview(false);
    // Add reload after dialog closes
    setTimeout(() => {
      window.location.reload();
    }, 300);
  };

  const PreviewDialog = () => (
    <Dialog
      fullScreen
      open={openPreview}
      onClose={handlePreviewClose}
      sx={{ '& .MuiDialog-paper': { bgcolor: '#f5f5f5' } }}
    >
      <AppBar sx={{ position: 'relative', bgcolor: 'white', color: 'black' }}>
        <Toolbar>
          <Typography sx={{ flex: 1 }} variant="h6">
            Excel Preview
          </Typography>
          <IconButton edge="end" color="inherit" onClick={handlePreviewClose}>
            <CloseIcon />
          </IconButton>
        </Toolbar>
      </AppBar>
      <DialogContent sx={{ p: 0 }}>
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
        <div className="excel-table-wrapper preview-mode">
          <table className="excel-table">
            <thead>
              <tr>
                <th className="row-header corner-header">#</th>
                {previewData?.headers.map((header, index) => (
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
              {previewData?.rows.map((row, rowIndex) => (
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
      </DialogContent>
    </Dialog>
  );

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
          <button 
            className="preview-button"
            onClick={handlePreviewOpen}
          >
            <PreviewIcon /> Preview File
          </button>
          <PreviewDialog />
        </>
      )}
    </div>
  );
};

export default ExcelImport;
