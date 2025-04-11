import React from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { Button, Typography, Box } from '@mui/material';
import FileDownloadIcon from '@mui/icons-material/FileDownload';

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

  const handleTitleClick = () => {
    window.location.reload();
  };

  return (
    <Box sx={{ textAlign: 'center' }}>
      <Typography 
        variant="h5" 
        gutterBottom
        sx={{ 
          cursor: 'pointer',
          color: '#1a73e8',
          '&:hover': { color: '#1557b0' }
        }} 
        onClick={handleTitleClick}
      >
        Excel Export Example
      </Typography>
      <Button
        variant="contained"
        startIcon={<FileDownloadIcon />}
        onClick={exportToExcel}
        sx={{
          backgroundColor: '#10ad4c',
          '&:hover': { backgroundColor: '#0d8c3e' }
        }}
      >
        Export to Excel
      </Button>
    </Box>
  );
};

export default ExcelExampleExport;
