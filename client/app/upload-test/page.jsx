"use client";
import { useState } from 'react';

export default function UploadTest() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [baseCode, setBaseCode] = useState('DENEMEKOD');
  const [startNumber, setStartNumber] = useState(0);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file1 || !file2) {
      alert("Please upload both files.");
      return;
    }

    const formData = new FormData();
    formData.append('file1', file1);
    formData.append('file2', file2);
    formData.append('baseCode', baseCode);
    formData.append('startNumber', startNumber);

    const response = await fetch('/api/process-excel', {
      method: 'POST',
      body: formData,
    });

    if (response.ok) {
      alert("File processed and saved successfully in the 'excels' directory.");
    } else {
      try {
        const errorData = await response.json();
        console.error('Error processing files:', errorData);
        alert(`Error: ${errorData.error}\nDetails: ${errorData.details}`);
      } catch (e) {
        console.error('Error processing files, but could not parse error details:', e);
      }
    }
  };
  
  return (
    <div>
      <h1>Test Excel Processing</h1>
      <form onSubmit={handleSubmit}>
        <label>
          Upload Excel File 1:
          <input type="file" onChange={(e) => setFile1(e.target.files[0])} required />
        </label>
        <br />
        <label>
          Upload Excel File 2:
          <input type="file" onChange={(e) => setFile2(e.target.files[0])} required />
        </label>
        <br />
        <label>
          Base Code:
          <input
            type="text"
            value={baseCode}
            onChange={(e) => setBaseCode(e.target.value)}
          />
        </label>
        <br />
        <label>
          Start Number:
          <input
            type="number"
            value={startNumber}
            onChange={(e) => setStartNumber(Number(e.target.value))}
          />
        </label>
        <br />
        <button type="submit">Process Excel Files</button>
      </form>
    </div>
  );
}
