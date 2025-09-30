// api/export.js
require('dotenv').config();
const express = require('express');
const serverless = require('serverless-http');
const axios = require('axios');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const app = express();

// Helper function to generate random dummy emails
function generateDummyEmail() {
  const domains = ['example.com', 'test.org', 'demo.net', 'sample.co', 'placeholder.io'];
  const firstNames = ['john', 'jane', 'alex', 'sarah', 'mike', 'lisa', 'david', 'emma', 'chris', 'anna'];
  const lastNames = ['smith', 'johnson', 'brown', 'davis', 'miller', 'wilson', 'moore', 'taylor', 'anderson', 'thomas'];
  
  const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
  const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
  const domain = domains[Math.floor(Math.random() * domains.length)];
  const randomNum = Math.floor(Math.random() * 999) + 1;
  
  return `${firstName}.${lastName}${randomNum}@${domain}`;
}

// Helper function to randomize email data
function randomizeEmailData(headers, rows) {
  // Find email column (case-insensitive)
  const emailColumnIndex = headers.findIndex(header => 
    header.toLowerCase().includes('email') || header.toLowerCase().includes('e-mail')
  );
  
  if (emailColumnIndex === -1) {
    // No email column found, return data unchanged
    return { headers, rows };
  }
  
  const emailColumnName = headers[emailColumnIndex];
  console.log(`Found email column: ${emailColumnName}, randomizing data...`);
  
  // Create a set to ensure unique dummy emails
  const usedEmails = new Set();
  
  // Randomize email data in rows
  const randomizedRows = rows.map(row => {
    if (row[emailColumnName] && row[emailColumnName].trim() !== '') {
      let dummyEmail;
      // Generate unique dummy email
      do {
        dummyEmail = generateDummyEmail();
      } while (usedEmails.has(dummyEmail));
      
      usedEmails.add(dummyEmail);
      
      return {
        ...row,
        [emailColumnName]: dummyEmail
      };
    }
    return row;
  });
  
  return { headers, rows: randomizedRows };
}

// Helper function to parse CSV data
function parseCSV(csvText) {
  const lines = csvText.trim().split('\n');
  const headers = lines[0].split(',').map(h => h.replace(/"/g, ''));
  const rows = lines.slice(1).map(line => {
    const values = line.split(',').map(v => v.replace(/"/g, ''));
    const row = {};
    headers.forEach((header, index) => {
      row[header] = values[index] || '';
    });
    return row;
  });
  
  // Randomize email data for privacy protection
  return randomizeEmailData(headers, rows);
}

// Helper function to convert to Excel
async function convertToExcel(csvData, filename) {
  const { headers, rows } = parseCSV(csvData);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');
  
  // Add headers
  worksheet.addRow(headers);
  
  // Style headers
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE6F3FF' }
  };
  
  // Add data rows
  rows.forEach(row => {
    worksheet.addRow(headers.map(header => row[header]));
  });
  
  // Auto-fit columns
  worksheet.columns.forEach(column => {
    column.width = Math.max(15, Math.min(50, column.header?.length || 15));
  });
  
  return await workbook.xlsx.writeBuffer();
}

// Helper function to convert to PDF
function convertToPDF(csvData, filename) {
  return new Promise((resolve, reject) => {
    try {
      const { headers, rows } = parseCSV(csvData);
      const doc = new PDFDocument({ margin: 50 });
      const buffers = [];
      
      doc.on('data', buffers.push.bind(buffers));
      doc.on('end', () => {
        const pdfData = Buffer.concat(buffers);
        resolve(pdfData);
      });
      
      // Title
      doc.fontSize(16).text(filename.replace(/\.[^/.]+$/, ""), 50, 50);
      doc.moveDown();
      
      // Table setup
      const cellWidth = 80;
      const cellHeight = 20;
      let currentY = doc.y;
      
      // Headers
      doc.fontSize(10).font('Helvetica-Bold');
      headers.forEach((header, index) => {
        doc.rect(50 + index * cellWidth, currentY, cellWidth, cellHeight)
           .fillAndStroke('#f0f0f0', '#000000');
        doc.fillColor('#000000')
           .text(header.substring(0, 12), 55 + index * cellWidth, currentY + 5, {
             width: cellWidth - 10,
             height: cellHeight - 10
           });
      });
      
      currentY += cellHeight;
      
      // Data rows (limit to fit on page)
      doc.font('Helvetica').fontSize(8);
      const maxRows = Math.min(rows.length, 30);
      
      for (let i = 0; i < maxRows; i++) {
        const row = rows[i];
        headers.forEach((header, index) => {
          doc.rect(50 + index * cellWidth, currentY, cellWidth, cellHeight)
             .stroke('#cccccc');
          doc.fillColor('#000000')
             .text(String(row[header] || '').substring(0, 15), 
                   55 + index * cellWidth, currentY + 5, {
                     width: cellWidth - 10,
                     height: cellHeight - 10
                   });
        });
        currentY += cellHeight;
        
        // Check if we need a new page
        if (currentY > 700) {
          doc.addPage();
          currentY = 50;
        }
      }
      
      if (rows.length > maxRows) {
        doc.moveDown().text(`... and ${rows.length - maxRows} more rows`);
      }
      
      doc.end();
    } catch (error) {
      reject(error);
    }
  });
}

// Helper function to convert to JSON
function convertToJSON(csvData) {
  const { rows } = parseCSV(csvData);
  return JSON.stringify({
    data: rows,
    exportedAt: new Date().toISOString(),
    totalRows: rows.length
  }, null, 2);
}

// Helper function to convert back to CSV format with randomized data
function convertToCSV(csvData) {
  const { headers, rows } = parseCSV(csvData);
  
  // Create CSV header row
  let csvOutput = headers.map(header => `"${header}"`).join(',') + '\n';
  
  // Add data rows
  rows.forEach(row => {
    const rowValues = headers.map(header => `"${row[header] || ''}"`);
    csvOutput += rowValues.join(',') + '\n';
  });
  
  return csvOutput;
}

app.get('/export', async (req, res) => {
  try {
    // 1. Read query params
    const { viewId, dashboard, user, format = 'csv' } = req.query;
    const ts = new Date().toISOString();

    // 2. Sign in to Tableau REST API
    const signinResp = await axios.post(
      `${process.env.TABLEAU_SERVER}/api/3.13/auth/signin`,
      {
        credentials: {
          personalAccessTokenName: process.env.PAT_NAME,
          personalAccessTokenSecret: process.env.PAT_SECRET,
          site: { contentUrl: process.env.SITE_CONTENT_URL }
        }
      }
    );
    const token = signinResp.data.credentials.token;
    const siteId = signinResp.data.credentials.site.id;

    // 3. Fetch CSV of the full view
    const dataResp = await axios.get(
      `${process.env.TABLEAU_SERVER}/api/3.13/sites/${siteId}/views/${viewId}/data?maxAge=0`,
      {
        responseType: 'text',
        headers: { 'X-Tableau-Auth': token }
      }
    );
    // 4. Sign out
    await axios.post(
      `${process.env.TABLEAU_SERVER}/api/3.13/auth/signout`,
      {},
      { headers: { 'X-Tableau-Auth': token } }
    );

    // 5. Append tracker row
    let csv = dataResp.data;
    csv += `"TRACKER","${user}","${dashboard}","${ts}"\n`;

    // 6. Handle different export formats
    const filename = `${dashboard}_${ts}`;
    
    switch (format.toLowerCase()) {
      case 'excel':
        try {
          const excelBuffer = await convertToExcel(csv, filename);
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}.xlsx"`);
          return res.send(excelBuffer);
        } catch (excelErr) {
          console.error('Excel conversion failed:', excelErr);
          return res.status(500).send('Excel export failed: ' + excelErr.message);
        }
        
      case 'pdf':
        try {
          const pdfBuffer = await convertToPDF(csv, filename);
          res.setHeader('Content-Type', 'application/pdf');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}.pdf"`);
          return res.send(pdfBuffer);
        } catch (pdfErr) {
          console.error('PDF conversion failed:', pdfErr);
          return res.status(500).send('PDF export failed: ' + pdfErr.message);
        }
        
      case 'json':
        try {
          const jsonData = convertToJSON(csv);
          res.setHeader('Content-Type', 'application/json');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}.json"`);
          return res.send(jsonData);
        } catch (jsonErr) {
          console.error('JSON conversion failed:', jsonErr);
          return res.status(500).send('JSON export failed: ' + jsonErr.message);
        }
        
      case 'csv':
      default:
        try {
          const randomizedCSV = convertToCSV(csv);
          res.setHeader('Content-Type', 'text/csv');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}.csv"`);
          return res.send(randomizedCSV);
        } catch (csvErr) {
          console.error('CSV conversion failed:', csvErr);
          return res.status(500).send('CSV export failed: ' + csvErr.message);
        }
    }

  } catch (err) {
    console.error(err);
    return res.status(500).send('Export failed: ' + err.message);
  }
});

module.exports.handler = serverless(app);
