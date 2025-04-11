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
        variant="h4" 
        gutterBottom
        sx={{ 
          cursor: 'pointer',
          background: 'linear-gradient(45deg, #1a73e8 30%, #2196F3 90%)',
          WebkitBackgroundClip: 'text',
          WebkitTextFillColor: 'transparent',
          fontWeight: 700,
          mb: 3,
          '&:hover': { 
            transform: 'scale(1.01)',
            transition: 'transform 0.2s ease-in-out'
          }
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
          background: 'linear-gradient(45deg, #2196F3 30%, #21CBF3 90%)',
          boxShadow: '0 3px 5px 2px rgba(33, 203, 243, .3)',
          color: 'white',
          padding: '10px 30px',
          '&:hover': {
            background: 'linear-gradient(45deg, #1a73e8 30%, #2196F3 90%)',
            transform: 'translateY(-1px)'
          }
        }}
      >
        Export to Excel
      </Button>
    </Box>
  );
};

export default ExcelExampleExport;
