import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Typography, Box, Paper, Button, Modal, IconButton } from '@mui/material';
import PreviewIcon from '@mui/icons-material/Visibility';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import CloseIcon from '@mui/icons-material/Close';
import { keyframes } from '@mui/system';
import LoadingSpinner from './components/LoadingSpinner';

const ExcelImport = ({ uploadHandler }) => {
  const [errorMessage, setErrorMessage] = useState('');
  const [previewData, setPreviewData] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);

  const handleFile = useCallback((file) => {
    setIsLoading(true);
    if (!file) {
      setErrorMessage('No file uploaded!');
      setIsLoading(false);
      return;
    }

    if (!(file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.type === 'application/vnd.ms-excel')) {
      setErrorMessage('Unknown file format. Only Excel files are supported.');
      setIsLoading(false);
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
      setIsLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, [uploadHandler]);

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

  const openFileInModal = () => setModalOpen(true);
  const closeFileModal = () => setModalOpen(false);

  const renderExcelPreview = () => {
    if (!previewData || !sheets.length) return null;
    const workbook = XLSX.read(previewData.file, { type: 'array' });

    return (
      <Box
        sx={{
          width: '95vw',
          maxWidth: '1400px',
          maxHeight: '85vh',
          bgcolor: 'background.paper',
          borderRadius: 3,
          boxShadow: 24,
          p: 3,
          position: 'relative',
          overflow: 'auto',
        }}
      >
        <IconButton
          onClick={closeFileModal}
          sx={{
            position: 'absolute',
            top: 8,
            right: 8,
            zIndex: 10,
            color: '#1a73e8',
            background: '#f5f5f5',
            '&:hover': { background: '#e3f2fd' }
          }}
        >
          <CloseIcon />
        </IconButton>
        <Box sx={{ display: 'flex', alignItems: 'center', mb: 2 }}>
          <PreviewIcon sx={{ color: '#1a73e8', mr: 1 }} />
          <Typography variant="h6" sx={{ color: '#1a73e8', fontWeight: 700 }}>
            Excel Preview
          </Typography>
        </Box>
        <Box sx={{ display: 'flex', gap: 1, mb: 2, overflowX: 'auto' }}>
          {sheets.map((sheet, idx) => (
            <Button
              key={sheet}
              variant={activeSheetIdx === idx ? "contained" : "outlined"}
              size="small"
              sx={{
                background: activeSheetIdx === idx ? 'linear-gradient(45deg, #1a73e8 30%, #2196F3 90%)' : '#f5f5f5',
                color: activeSheetIdx === idx ? 'white' : '#1a73e8',
                fontWeight: 600,
                borderRadius: 2,
                minWidth: 100,
                border: activeSheetIdx === idx ? 'none' : '1px solid #e3f2fd',
                boxShadow: activeSheetIdx === idx ? 2 : 0,
                '&:hover': {
                  background: activeSheetIdx === idx
                    ? 'linear-gradient(45deg, #2196F3 30%, #21CBF3 90%)'
                    : '#e3f2fd'
                }
              }}
              onClick={() => setActiveSheetIdx(idx)}
            >
              {sheet}
            </Button>
          ))}
        </Box>
        <div className="excel-table-wrapper preview-mode" style={{ maxHeight: '60vh', overflow: 'auto' }}>
          {(() => {
            const ws = workbook.Sheets[sheets[activeSheetIdx]];
            const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            const headers = jsonData[0] || [];
            const rows = jsonData.slice(1);
            const colLetters = headers.map((_, i) => String.fromCharCode(65 + i));
            return (
              <table className="excel-table">
                <thead>
                  <tr>
                    <th className="row-header corner-header" style={{ background: '#f8f9fa', zIndex: 3, left: 0, borderRight: '2px solid #ddd', borderBottom: '2px solid #ddd' }}>#</th>
                    {headers.map((header, i) => (
                      <th className="column-cell" key={i} style={{ background: '#f5f6fa', position: 'sticky', top: 0, zIndex: 2 }}>
                        <div className="column-letter" style={{ color: '#666', fontSize: 12, fontWeight: 'normal', borderBottom: '1px dashed #ddd', marginBottom: 4, padding: '2px 0' }}>
                          {colLetters[i]}
                        </div>
                        <div className="header-content" style={{ fontWeight: 'bold', color: '#333' }}>
                          {header}
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, rowIndex) => (
                    <tr key={rowIndex} style={{ background: rowIndex === 0 ? '#3d5662' : rowIndex % 2 === 0 ? '#f9f9f9' : '#c6e6f5', color: rowIndex === 0 ? '#fff' : undefined }}>
                      <td className="row-header" style={{ position: 'sticky', left: 0, zIndex: 1, background: '#f5f6fa', color: '#666', fontSize: 12, fontWeight: 'normal', textAlign: 'center', minWidth: 30, borderRight: '2px solid #ddd' }}>
                        {rowIndex + 1}
                      </td>
                      {row.map((cell, cellIndex) => (
                        <td key={cellIndex} style={{ padding: '5px 10px' }}>{cell}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            );
          })()}
        </div>
      </Box>
    );
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

  return (
    <Box className="excel-import-container" sx={{ p: 3 }}>
      {isLoading && <LoadingSpinner />}
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
      {errorMessage && <p style={{ color: 'red' }}>{errorMessage}</p>}
      {previewData && (
        <Box sx={{ display: 'flex', justifyContent: 'center', mt: 2 }}>
          <Button
            variant="contained"
            startIcon={<PreviewIcon />}
            onClick={openFileInModal}
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
            View File
          </Button>
          <Modal
            open={modalOpen}
            onClose={closeFileModal}
            aria-labelledby="excel-preview-modal"
            sx={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 2000,
              background: 'rgba(0,0,0,0.25)'
            }}
          >
            {renderExcelPreview()}
          </Modal>
        </Box>
      )}
    </Box>
  );
};

export default ExcelImport;
