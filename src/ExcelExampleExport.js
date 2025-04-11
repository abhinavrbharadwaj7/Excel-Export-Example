import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { Fab } from '@mui/material';
import FileDownloadIcon from '@mui/icons-material/FileDownload';
import LoadingSpinner from './components/LoadingSpinner';

const ExcelExampleExport = () => {
  const [isLoading, setIsLoading] = useState(false);

  const data = [
    { name: 'John', age: 30, city: 'New York' },
    { name: 'Jane', age: 25, city: 'San Francisco' },
  ];

  const exportToExcel = () => {
    setIsLoading(true);
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const dataBlob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(dataBlob, 'example.xlsx');
    
    // Add page reload after a short delay to ensure file download completes
    setTimeout(() => {
      setIsLoading(false);
      window.location.reload();
    }, 1000);
  };

  return (
    <>
      {isLoading && <LoadingSpinner />}
      <Fab
        color="primary"
        aria-label="export"
        onClick={exportToExcel}
        sx={{
          position: 'fixed',
          bottom: 20,
          right: 20,
          background: 'linear-gradient(45deg, #2196F3 30%, #21CBF3 90%)',
          boxShadow: '0 3px 5px 2px rgba(33, 150, 243, .3)',
          '&:hover': {
            background: 'linear-gradient(45deg, #1a73e8 30%, #2196F3 90%)',
            transform: 'scale(1.1)',
            transition: 'all 0.3s ease'
          }
        }}
      >
        <FileDownloadIcon />
      </Fab>
    </>
  );
};

export default ExcelExampleExport;
