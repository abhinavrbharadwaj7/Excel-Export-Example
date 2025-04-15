import React from 'react';
import { Box, keyframes } from '@mui/material';
import TableChartIcon from '@mui/icons-material/TableChart';

const rotate = keyframes`
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
`;

const wave = keyframes`
  0% { transform: translateY(0); }
  50% { transform: translateY(-10px); }
  100% { transform: translateY(0); }
`;

const LoadingSpinner = () => (
  <Box
    sx={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      background: 'rgba(255, 255, 255, 0.9)',
      backdropFilter: 'blur(8px)',
      zIndex: 9999,
    }}
  >
    <Box
      sx={{
        position: 'relative',
        width: '80px',
        height: '80px',
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
      }}
    >
      <Box
        sx={{
          position: 'absolute',
          width: '60px',
          height: '60px',
          border: '4px solid transparent',
          borderTop: '4px solid #1a73e8',
          borderRadius: '50%',
          animation: `${rotate} 1s linear infinite`,
        }}
      />
      <TableChartIcon
        sx={{
          fontSize: '30px',
          color: '#1a73e8',
          animation: `${wave} 1s ease-in-out infinite`,
        }}
      />
    </Box>
    <Box
      sx={{
        marginTop: '20px',
        color: '#1a73e8',
        fontWeight: 500,
        animation: `${wave} 1s ease-in-out infinite`,
        animationDelay: '0.1s',
      }}
    >
      Loading...
    </Box>
  </Box>
);

export default LoadingSpinner;
