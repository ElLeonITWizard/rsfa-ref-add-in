import * as React from 'react';
import { Box, Paper, Typography, IconButton } from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';

function DragDropFileUpload({ onFileUpload }) {
  const [dragOver, setDragOver] = React.useState(false);

  const handleDragOver = React.useCallback((event) => {
    event.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = React.useCallback((event) => {
    event.preventDefault();
    setDragOver(false);
  }, []);

  const handleDrop = React.useCallback(
    (event) => {
      event.preventDefault();
      setDragOver(false);
      if (event.dataTransfer.files && event.dataTransfer.files[0]) {
        onFileUpload(event.dataTransfer.files[0]);
      }
    },
    [onFileUpload]
  );

  const handleChange = React.useCallback(
    (event) => {
      if (event.target.files && event.target.files[0]) {
        onFileUpload(event.target.files[0]);
      }
    },
    [onFileUpload]
  );

  return (
    <Paper
      variant="outlined"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      style={{
        border: dragOver ? '2px dashed #000' : '2px dashed #aaa',
        padding: 20,
        textAlign: 'center',
        cursor: 'pointer',
        background: dragOver ? '#eee' : '#fafafa',
      }}
    >
      <input
        //accept="image/*"
        style={{ display: 'none' }}
        id="raised-button-file"
        multiple
        type="file"
        onChange={handleChange}
      />
      <label htmlFor="raised-button-file">
        <Box display="flex" flexDirection="column" alignItems="center">
          <IconButton color="primary" aria-label="upload picture" component="span">
            <CloudUploadIcon style={{ fontSize: 60 }} />
          </IconButton>
          <Typography>Drag and drop files here or click to select files</Typography>
        </Box>
      </label>
    </Paper>
  );
}

export default DragDropFileUpload;