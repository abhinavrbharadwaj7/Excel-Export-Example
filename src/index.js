import React, { useState } from "react";
import ReactDOM from "react-dom";
import { Box, Container, AppBar, Toolbar, Typography, Paper, Stack } from '@mui/material';
import TableChartIcon from '@mui/icons-material/TableChart';
import ExcelImport from "./ExcelImport";
import ExcelExampleExport from "./ExcelExampleExport";
import "./styles.css";

const App = () => {
  const [data, setData] = useState([]);

  const createRequests = () => {
    console.log(data);
  };

  const handleTitleClick = () => {
    window.location.reload();
  };

  const worksheets = [
    {
      name: "Requests",
      columns: [
        { label: "Full Name", value: "name" },
        { label: "Email", value: "email" },
        { label: "Template", value: "template" }
      ],
      data: [
        {
          name: "Bob Ross",
          email: "boss_ross@gmail.com",
          template: "Accounts Receivables"
        }
      ]
    }
  ];

  return (
    <Box sx={{ flexGrow: 1 }}>
      <AppBar 
        position="fixed" 
        sx={{ 
          background: 'rgba(255, 255, 255, 0.7)',
          backdropFilter: 'blur(10px)',
          boxShadow: '0 4px 30px rgba(0, 0, 0, 0.1)',
          borderBottom: '1px solid rgba(255, 255, 255, 0.3)',
          marginBottom: 4 
        }}
      >
        <Toolbar>
          <TableChartIcon sx={{ mr: 2, color: '#1a73e8' }} />
          <Typography 
            variant="h6" 
            component="div" 
            sx={{ 
              flexGrow: 1, 
              cursor: 'pointer',
              color: '#1a73e8',
              fontWeight: 600,
              '&:hover': { 
                opacity: 0.8,
                transform: 'scale(1.01)',
                transition: 'all 0.2s ease-in-out'
              }
            }}
            onClick={handleTitleClick}
          >
            Excel Processor
          </Typography>
        </Toolbar>
      </AppBar>
      <Toolbar /> {/* Spacer for fixed AppBar */}
      
      <Container maxWidth="md" sx={{ mt: 4 }}>
        <Stack spacing={4}>
          <Paper 
            elevation={0} 
            sx={{ 
              p: 3, 
              backgroundColor: '#f8f9fa',
              borderRadius: '16px',
              transition: 'transform 0.2s ease-in-out',
              '&:hover': {
                transform: 'scale(1.01)'
              }
            }}
          >
            <ExcelExampleExport filename="requests.xlsx" worksheets={worksheets} />
          </Paper>

          <Paper 
            elevation={2} 
            sx={{ 
              p: 0, 
              overflow: 'hidden',
              borderRadius: '16px'
            }}
          >
            <ExcelImport uploadHandler={setData} />
          </Paper>
        </Stack>
      </Container>
    </Box>
  );
};

const rootElement = document.getElementById("root");
ReactDOM.render(<App />, rootElement);
