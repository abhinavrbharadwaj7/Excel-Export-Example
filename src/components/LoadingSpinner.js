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

const glow = keyframes`
  0%, 100% { filter: drop-shadow(0 0 8px rgba(26, 115, 232, 0.6)); }
  50% { filter: drop-shadow(0 0 15px rgba(26, 115, 232, 0.8)); }
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
      background: 'rgba(255, 255, 255, 0.95)',
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
      {/* Single rotating ring */}
      <Box
        sx={{
          position: 'absolute',
          width: '100%',
          height: '100%',
          border: '3px solid transparent',
          borderTop: '3px solid #1a73e8',
          borderRight: '3px solid #1a73e8',
          borderRadius: '50%',
          animation: `${rotate} 1.5s linear infinite`,
        }}
      />

      {/* Center icon */}
      <TableChartIcon
        sx={{
          fontSize: '35px',
          color: '#1a73e8',
          animation: `${wave} 1.5s ease-in-out infinite, ${glow} 2s ease-in-out infinite`,
        }}
      />
    </Box>

    <Box
      sx={{
        marginTop: '20px',
        fontSize: '16px',
        fontWeight: 500,
        background: 'linear-gradient(45deg, #1a73e8, #2196F3)',
        WebkitBackgroundClip: 'text',
        WebkitTextFillColor: 'transparent',
        animation: `${wave} 1.5s ease-in-out infinite`,
      }}
    >
      Processing...
    </Box>
  </Box>
);

export default LoadingSpinner;
