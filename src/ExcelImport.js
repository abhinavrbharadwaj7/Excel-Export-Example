import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Dialog, DialogContent, DialogTitle, IconButton, AppBar, Toolbar, Typography, Box, Paper, Button } from '@mui/material';
import CloseIcon from '@mui/icons-material/Close';
import PreviewIcon from '@mui/icons-material/Visibility';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import { keyframes } from '@mui/system';

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

  const rippleAnimation = keyframes`
    0% {
      transform: scale(0.95);
      box-shadow: 0 0 0 0 rgba(26, 115, 232, 0.3);
    }
    70% {
      transform: scale(1);
      box-shadow: 0 0 0 10px rgba(26, 115, 232, 0);
    }
    100% {
      transform: scale(0.95);
      box-shadow: 0 0 0 0 rgba(26, 115, 232, 0);
    }
  `;

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
    <Box className="excel-import-container" sx={{ p: 3 }}>
      <Typography 
        variant="h4" 
        gutterBottom
        sx={{ 
          background: 'linear-gradient(45deg, #1a73e8 30%, #2196F3 90%)',
          WebkitBackgroundClip: 'text',
          WebkitTextFillColor: 'transparent',
          fontWeight: 700,
          mb: 3,
          textAlign: 'center'
        }}
      >
        Import Excel File
      </Typography>
      
      <Paper
        elevation={0}
        className={`file-upload-area ${isDragging ? 'dragging' : ''}`}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        sx={{
          border: '2px dashed #1a73e8',
          borderRadius: 2,
          p: 3,
          textAlign: 'center',
          backgroundColor: isDragging ? 'rgba(26, 115, 232, 0.1)' : 'transparent',
          transition: 'all 0.3s ease',
          minHeight: '200px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center'
        }}
      >
        <Box className="file-upload">
          <input
            type="file"
            onChange={fileHandler}
            accept=".xlsx, .xls"
            style={{ display: 'none' }}
            id="excel-file-input"
          />
          <label htmlFor="excel-file-input" style={{ width: '100%' }}>
            <Box sx={{ 
              display: 'flex', 
              flexDirection: 'column', 
              alignItems: 'center',
              justifyContent: 'center',
              gap: 2,
              cursor: 'pointer'
            }}>
              <CloudUploadIcon sx={{ fontSize: 48, color: '#1a73e8' }} />
              <Typography variant="body1" color="textSecondary">
                Drag and drop Excel file here or click to browse
              </Typography>
            </Box>
          </label>
        </Box>
      </Paper>

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
        <Box sx={{ display: 'flex', justifyContent: 'center', mt: 2 }}>
          <Button
            variant="contained"
            startIcon={<PreviewIcon />}
            onClick={handlePreviewOpen}
            sx={{
              background: 'linear-gradient(45deg, #1a73e8 30%, #2196F3 90%)',
              boxShadow: '0 3px 5px 2px rgba(33, 150, 243, .3)',
              color: 'white',
              padding: '10px 30px',
              borderRadius: '25px',
              fontWeight: 600,
              transition: 'all 0.3s ease',
              animation: `${rippleAnimation} 1.5s infinite`,
              '&:hover': {
                background: 'linear-gradient(45deg, #2196F3 30%, #21CBF3 90%)',
                transform: 'translateY(-2px)',
                boxShadow: '0 6px 10px 2px rgba(33, 150, 243, .3)',
              },
              '&:active': {
                transform: 'translateY(1px)',
              }
            }}
          >
            Preview File
          </Button>
          <PreviewDialog />
        </Box>
      )}
    </Box>
  );
};

export default ExcelImport;
