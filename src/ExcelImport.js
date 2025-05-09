import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Typography, Box, Paper, Button } from '@mui/material';
import PreviewIcon from '@mui/icons-material/Visibility';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import { keyframes } from '@mui/system';
import LoadingSpinner from './components/LoadingSpinner';

const ExcelImport = ({ uploadHandler }) => {
  const [errorMessage, setErrorMessage] = useState('');
  const [previewData, setPreviewData] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  const handleFile = useCallback((file) => {
    setIsLoading(true);
    setTimeout(() => { // Ensure spinner renders before heavy work
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
          // Ensure all rows have the correct length
          let rows = jsonData.slice(1).map(row => {
            const newRow = new Array(headers.length).fill('');
            row.forEach((cell, idx) => {
              newRow[idx] = cell;
            });
            return newRow;
          });

          // Handle merged cells
          const mergedCells = ws['!merges'] || [];
          mergedCells.forEach((merge) => {
            const startRow = merge.s.r;
            const startCol = merge.s.c;
            const endRow = merge.e.r;
            const endCol = merge.e.c;
            // Adjust for header row
            const value = rows[startRow]?.[startCol] || '';
            for (let r = startRow; r <= endRow; r++) {
              if (!rows[r]) rows[r] = new Array(headers.length).fill('');
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
    }, 0);
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

  const openFileInNewTab = () => {
    const newWindow = window.open('', '_blank');
    const workbook = XLSX.read(previewData.file, { type: 'array' });
    const sheetNames = workbook.SheetNames;

    let sheetTabs = '';
    let sheetContents = '';
    sheetNames.forEach((sheet, idx) => {
      const ws = workbook.Sheets[sheet];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const headers = jsonData[0] || [];
      const rows = jsonData.slice(1);
      const colLetters = headers.map((_, i) => String.fromCharCode(65 + i));

      sheetTabs += `<button class="sheet-tab" id="tab-${idx}" onclick="showSheet(${idx})" style="margin-right:8px;${idx === 0 ? 'font-weight:bold;background:#e3f2fd;' : ''}">${sheet}</button>`;

      sheetContents += `
        <div id="sheet-content-${idx}" style="display:${idx === 0 ? 'block' : 'none'};">
          <div class="excel-table-wrapper preview-mode" style="height:600px;max-height:600px;overflow:auto;overflow-x:auto;">
            <table class="excel-table" style="min-width:100%;width:max-content;">
              <thead>
                <tr>
                  <th class="row-header corner-header" style="background:#f8f9fa;z-index:3;left:0;border-right:2px solid #ddd;border-bottom:2px solid #ddd;">#</th>
                  ${headers.map((header, i) => `
                    <th class="column-cell" style="background:#f5f6fa;position:sticky;top:0;z-index:2;">
                      <div class="column-letter" style="color:#666;font-size:12px;font-weight:normal;border-bottom:1px dashed #ddd;margin-bottom:4px;padding:2px 0;">
                        ${colLetters[i]}
                      </div>
                      <div class="header-content" style="font-weight:bold;color:#333;">
                        ${header}
                      </div>
                    </th>
                  `).join('')}
                </tr>
              </thead>
              <tbody>
                ${rows.map((row, rowIndex) => `
                  <tr style="background:${rowIndex === 0 ? '#3d5662;color:#fff;' : rowIndex % 2 === 0 ? '#f9f9f9;' : '#c6e6f5;'}">
                    <td class="row-header" style="position:sticky;left:0;z-index:1;background:#f5f6fa;color:#666;font-size:12px;font-weight:normal;text-align:center;min-width:30px;border-right:2px solid #ddd;">
                      ${rowIndex + 1}
                    </td>
                    ${row.map((cell, cellIndex) => `
                      <td style="padding:5px 10px;">${cell}</td>
                    `).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    });

    newWindow.document.write(`
      <html>
        <head>
          <title>Excel Preview</title>
          <style>
            html, body { 
              height: 100%; 
              margin: 0; 
              padding: 0; 
              font-family: Arial, sans-serif;
              background: #fff;
            }
            .header {
              position: fixed;
              top: 0;
              left: 0;
              right: 0;
              background: white;
              padding: 8px 16px;
              display: flex;
              align-items: center;
              gap: 12px;
              border-bottom: 1px solid #ddd;
              z-index: 1000;
            }
            .content {
              padding-top: 50px; /* header height */
              height: calc(100vh - 50px);
              overflow: auto;
            }
            .sheet-tabs-bar {
              position: sticky;
              top: 0;
              background: white;
              padding: 8px 16px 0;
              border-bottom: 1px solid #ddd;
              z-index: 100;
            }
            .sheet-tab {
              padding: 8px 16px;
              border: none;
              border-radius: 4px 4px 0 0;
              background: #f5f5f5;
              cursor: pointer;
              margin-right: 8px;
            }
            .sheet-tab.active, .sheet-tab:focus {
              font-weight: bold;
              background: #e3f2fd;
              outline: none;
            }
            .excel-table-wrapper {
              margin: 0;
              overflow: auto;
              border: none;
            }
            .excel-table {
              width: max-content;
              min-width: 100%;
              border-collapse: separate;
              border-spacing: 0;
            }
            .excel-table th, .excel-table td { padding: 8px; border: 1px solid #ddd; text-align: left; }
            .excel-table th { background-color: #f5f6fa; position: sticky; top: 0; z-index: 2; }
            .excel-table .row-header { position: sticky; left: 0; z-index: 1; background-color: #f5f6fa; color: #666; font-size: 12px; font-weight: normal; text-align: center; min-width: 30px; border-right: 2px solid #ddd; }
            .excel-table tr:first-of-type { display: table-row; }
            .excel-table tr { background-color: #c6e6f5; }
            .excel-table tr:nth-of-type(2) { background-color: #3d5662 !important; color: #fff; }
            .excel-table tr:nth-of-type(even) { background-color: #f9f9f9; }
            .excel-table tr:hover { background-color: #f5f5f5; }
            .excel-table td { padding: 5px 10px; }
            .excel-table .column-cell { padding: 4px 8px; position: relative; }
            .excel-table .column-letter { color: #666; font-size: 12px; font-weight: normal; border-bottom: 1px dashed #ddd; margin-bottom: 4px; padding: 2px 0; }
            .excel-table .header-content { font-weight: bold; color: #333; }
            .excel-table .corner-header { z-index: 3; left: 0; background-color: #f8f9fa; border-right: 2px solid #ddd; border-bottom: 2px solid #ddd; }
            .zoom-control {
              position: fixed;
              bottom: 20px;
              right: 20px;
              background: white;
              padding: 8px;
              border-radius: 20px;
              box-shadow: 0 2px 8px rgba(0,0,0,0.15);
              display: flex;
              align-items: center;
              gap: 8px;
              z-index: 1000;
            }
            .zoom-slider {
              width: 100px;
              height: 4px;
              -webkit-appearance: none;
              background: #e0e0e0;
              border-radius: 2px;
              outline: none;
            }
            .zoom-slider::-webkit-slider-thumb {
              -webkit-appearance: none;
              width: 16px;
              height: 16px;
              background: #1a73e8;
              border-radius: 50%;
              cursor: pointer;
              transition: all 0.2s ease;
            }
            .zoom-slider::-webkit-slider-thumb:hover {
              transform: scale(1.2);
            }
            .zoom-value {
              min-width: 45px;
              color: #666;
              font-size: 13px;
              text-align: right;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <svg style="width:24px;height:24px;color:#1a73e8" viewBox="0 0 24 24">
              <path fill="currentColor" d="M19,3H5C3.89,3 3,3.89 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5C21,3.89 20.1,3 19,3M19,19H5V5H19V19M17,17H7V7H17V17Z" />
            </svg>
            <span style="color:#1a73e8;font-size:18px;font-weight:600;">Excel Processor</span>
          </div>
          <div class="content">
            <div class="sheet-tabs-bar">${sheetTabs}</div>
            ${sheetContents}
          </div>
          <div class="zoom-control">
            <input type="range" class="zoom-slider" min="50" max="200" value="100" id="zoomSlider">
            <span class="zoom-value" id="zoomValue">100%</span>
          </div>
          <script>
            function showSheet(idx) {
              var total = ${sheetNames.length};
              for (var i = 0; i < total; i++) {
                document.getElementById('sheet-content-' + i).style.display = (i === idx) ? 'block' : 'none';
                document.getElementById('tab-' + i).style.fontWeight = (i === idx) ? 'bold' : 'normal';
                document.getElementById('tab-' + i).style.background = (i === idx) ? '#e3f2fd' : '#f5f5f5';
              }
            }

            const zoomSlider = document.getElementById('zoomSlider');
            const zoomValue = document.getElementById('zoomValue');
            const tables = document.querySelectorAll('.excel-table');
            
            zoomSlider.addEventListener('input', function() {
              const zoom = this.value;
              zoomValue.textContent = zoom + '%';
              tables.forEach(table => {
                table.style.transform = 'scale(' + (zoom/100) + ')';
                table.style.transformOrigin = 'top left';
              });
            });
          </script>
        </body>
      </html>
    `);
    newWindow.document.close();
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
            onClick={openFileInNewTab}
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
        </Box>
      )}
    </Box>
  );
};

export default ExcelImport;
