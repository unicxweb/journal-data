const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const app = express();
const port = 3000;

// Ensure 'uploads' and 'public' folders exist in working directory
const uploadDir = path.join(process.cwd(), 'uploads');
const publicDir = path.join(process.cwd(), 'public');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir);

// Serve static files
app.use(express.static(publicDir));
app.use('/downloads', express.static(uploadDir));

// Home route
app.get('/', (req, res) => {
  res.send('✅ Server is running! Use POST /upload to upload an Excel file.');
});

// Multer config
const upload = multer({ dest: uploadDir });

// API Config
const API_KEY = 'kEgyMhuf.sHUq798cDnqCS8C75mvSH1hUd57Gm1at';
const API_URL = 'https://api.binbash.ai/api/v2/trademarks/';

// Upload route
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send({ message: 'No file uploaded.' });

    const filePath = path.join(uploadDir, req.file.filename);

    // Validate file extension
    if (!req.file.originalname.match(/\.(xls|xlsx)$/)) {
      fs.unlinkSync(filePath);
      return res.status(400).send({ message: 'Only Excel files are allowed.' });
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    const updatedData = [];

    console.log('🔄 Processing Excel rows...');

    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      const rowNum = i + 2;
      const applicationNumber = row['Application Number'];
      const currentStatus = row['Status'];

      if (applicationNumber && currentStatus) {
        console.log(`➡️ Row ${rowNum}: Checking ${applicationNumber}`);
        const newStatus = await getStatusFromAPI(applicationNumber);

        if (newStatus && newStatus !== currentStatus) {
          console.log(`✅ Row ${rowNum}: Updated status: ${currentStatus} → ${newStatus}`);
          row['Status'] = newStatus;
        } else {
          console.log(`ℹ️ Row ${rowNum}: No change in status`);
        }
      } else {
        console.warn(`⚠️ Row ${rowNum}: Missing 'Application Number' or 'Status'`);
      }

      updatedData.push(row);
    }

    const updatedWorksheet = xlsx.utils.json_to_sheet(updatedData);
    const updatedWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, 'Updated Data');

    const updatedFileName = `updated_${Date.now()}.xlsx`;
    const updatedFilePath = path.join(uploadDir, updatedFileName);
    const publicDownloadPath = `/downloads/${updatedFileName}`;

    xlsx.writeFile(updatedWorkbook, updatedFilePath);
    fs.unlinkSync(filePath);

    console.log('✅ Excel processing completed.');
    res.json({ message: 'File processed successfully.', updatedFile: publicDownloadPath });

  } catch (err) {
    console.error('🚨 Processing error:', err);
    res.status(500).send({ message: 'Error processing file.' });
  }
});

// Function to fetch status from the API using Node 18 global fetch
async function getStatusFromAPI(applicationNumber) {
  const url = `${API_URL}${applicationNumber}`;
  const options = { method: 'GET', headers: { Authorization: `Api-Key ${API_KEY}` } };

  try {
    console.log(`🌐 Fetching API for ${applicationNumber}`);
    const res = await fetch(url, options);
    const text = await res.text();

    if (text.startsWith('<!doctype html>')) throw new Error('HTML error page received');
    const data = JSON.parse(text);
    return data.status || null;

  } catch (err) {
    console.error(`❌ API fetch error for ${applicationNumber}:`, err);
    return null;
  }
}

// Start the server
app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
  exec(`start http://localhost:${port}`);
});
