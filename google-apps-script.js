/**
 * Enhanced Google Apps Script for Tableau Multi-Format Export - VERSION 3.1
 * Deploy as a Web App (Execute as: Me / Anyone with access).
 * This version receives data via POST, simplifying the backend logic.
 * Includes a doGet for health checks.
 */

const SCRIPT_VERSION = '3.3-ULTIMATE-DEBUG-VERSION';

// Email notification settings
const SEND_EMAIL_NOTIFICATIONS = true;
const NOTIFICATION_EMAIL = 'gert.verbeken@easyfairs.com';

/**
 * Generate random dummy emails for privacy protection
 */
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

/**
 * Randomize email data for privacy protection
 */
function randomizeEmailData(csvData) {
  console.log('🔍🔍🔍 STARTING EMAIL RANDOMIZATION PROCESS 🔍🔍🔍');
  console.log('Raw CSV data length:', csvData.length);
  console.log('First 200 chars of CSV:', csvData.substring(0, 200));
  
  const lines = csvData.split('\n').filter(line => line.trim().length > 0);
  console.log('Total lines after filtering:', lines.length);
  
  if (lines.length < 2) {
    console.log('❌ Not enough lines for processing');
    return csvData;
  }
  
  const headers = parseCSVLine(lines[0]);
  console.log('🏷️ EXACT HEADERS FOUND:');
  headers.forEach((header, index) => {
    console.log(`  [${index}] "${header}" (length: ${header.length})`);
  });
  
  // Enhanced email column detection with more patterns
  let emailColumnIndex = -1;
  let matchedPattern = '';
  
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const lowerHeader = header.toLowerCase().trim();
    
    // Check various email patterns
    if (lowerHeader === 'email' || 
        lowerHeader === 'e-mail' || 
        lowerHeader === 'mail' ||
        lowerHeader === 'email address' ||
        lowerHeader === 'e-mail address' ||
        lowerHeader === 'emailaddress' ||
        lowerHeader.includes('email') || 
        lowerHeader.includes('e-mail') || 
        lowerHeader.includes('mail')) {
      emailColumnIndex = i;
      matchedPattern = lowerHeader;
      break;
    }
  }
  
  if (emailColumnIndex === -1) {
    console.log('❌❌❌ NO EMAIL COLUMN FOUND! ❌❌❌');
    console.log('Searched patterns: email, e-mail, mail, email address, etc.');
    console.log('Available headers again:');
    headers.forEach((header, index) => {
      console.log(`  [${index}] "${header}"`);
    });
    return csvData;
  }
  
  const emailColumnName = headers[emailColumnIndex];
  console.log(`✅✅✅ FOUND EMAIL COLUMN: "${emailColumnName}" at index ${emailColumnIndex}`);
  console.log(`🎯 Matched pattern: "${matchedPattern}"`);
  
  const usedEmails = new Set();
  let emailsRandomized = 0;
  
  // Process data rows (skip header)
  const processedLines = [lines[0]]; // Keep header unchanged
  
  for (let i = 1; i < lines.length; i++) {
    console.log(`🔄 Processing row ${i}: "${lines[i]}"`);
    const cells = parseCSVLine(lines[i]);
    console.log(`🔄 Parsed cells for row ${i}:`, cells);
    console.log(`🔄 Email cell [${emailColumnIndex}]: "${cells[emailColumnIndex]}"`);
    
    if (cells[emailColumnIndex] && cells[emailColumnIndex].trim() !== '') {
      const originalEmail = cells[emailColumnIndex];
      let dummyEmail;
      do {
        dummyEmail = generateDummyEmail();
      } while (usedEmails.has(dummyEmail));
      
      usedEmails.add(dummyEmail);
      cells[emailColumnIndex] = dummyEmail;
      emailsRandomized++;
      
      console.log(`✅ Row ${i}: "${originalEmail}" -> "${dummyEmail}"`);
      console.log(`✅ Updated cells:`, cells);
    } else {
      console.log(`⏭️ Row ${i}: No email to replace (empty or null)`);
    }
    
    // Rebuild line with proper CSV escaping
    const escapedCells = cells.map(cell => `"${cell.replace(/"/g, '""')}"`);
    const rebuiltLine = escapedCells.join(',');
    console.log(`🔄 Rebuilt line ${i}: "${rebuiltLine}"`);
    processedLines.push(rebuiltLine);
  }
  
  console.log(`🎯 Email randomization complete: ${emailsRandomized} emails replaced`);
  
  const finalCsv = processedLines.join('\n');
  console.log('📤 FINAL RANDOMIZED CSV (first 300 chars):');
  console.log(finalCsv.substring(0, 300));
  console.log('📤 Final CSV total length:', finalCsv.length);
  
  return finalCsv;
}

/**
 * Handles GET requests - provides a status message.
 */
function doGet(e) {
  return ContentService.createTextOutput(`Tableau Export Backend v${SCRIPT_VERSION} is running. Use POST to send data.`)
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Main function - handles POST requests with data from the extension.
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    
    const {
      csvData,
      user = 'unknown',
      dashboard = 'unknown',
      format = 'csv'
    } = postData;

    if (!csvData) {
      throw new Error('No CSV data received.');
    }

    // 🔒 PRIVACY PROTECTION: Randomize email data before processing
    const emailRandomizedCsv = randomizeEmailData(csvData);
    
    const ts = new Date().toISOString();
    const trackerCode = 'TRACKER-' + generateUniqueId();
    
    // Append tracker information to the CSV data
    const finalCsv = emailRandomizedCsv + '\n"' + trackerCode + '","' + user + '","' + dashboard + '","' + ts + '"';
    
    const safeFilename = dashboard.replace(/\W+/g, '_') + '_' + ts.replace(/[:.]/g, '-');

    switch (format.toLowerCase()) {
      case 'excel':
        return handleExcelExport(finalCsv, safeFilename, trackerCode, user, dashboard);
      case 'pdf':
        return handlePDFExport(finalCsv, safeFilename, trackerCode, user, dashboard);
      case 'json':
        return handleJSONExport(finalCsv, safeFilename, trackerCode, user, dashboard);
      default:
        return handleCSVExport(finalCsv, safeFilename, trackerCode, user, dashboard);
    }
    
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Handle CSV export
 */
function handleCSVExport(csv, safeFilename, trackerCode, user, dashboard) {
  if (SEND_EMAIL_NOTIFICATIONS) {
    sendExportNotification(user, dashboard, csv.split('\n').length - 2, trackerCode, 'CSV');
  }
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Handle Excel export
 */
function handleExcelExport(csv, safeFilename, trackerCode, user, dashboard) {
  const htmlTable = convertCSVToHTMLTable(csv, dashboard);
  if (SEND_EMAIL_NOTIFICATIONS) {
    sendExportNotification(user, dashboard, csv.split('\n').length - 2, trackerCode, 'Excel');
  }
  return ContentService.createTextOutput(htmlTable).setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Handle PDF export
 */
function handlePDFExport(csv, safeFilename, trackerCode, user, dashboard) {
  const htmlContent = convertCSVToPDF(csv, dashboard);
  if (SEND_EMAIL_NOTIFICATIONS) {
    sendExportNotification(user, dashboard, csv.split('\n').length - 2, trackerCode, 'PDF');
  }
  const pdfBlob = Utilities.newBlob(htmlContent, 'text/html').getAs('application/pdf');
  return ContentService.createTextOutput(Utilities.base64Encode(pdfBlob.getBytes())).setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Handle JSON export
 */
function handleJSONExport(csv, safeFilename, trackerCode, user, dashboard) {
  const jsonData = convertCSVToJSON(csv, dashboard);
  if (SEND_EMAIL_NOTIFICATIONS) {
    sendExportNotification(user, dashboard, csv.split('\n').length - 2, trackerCode, 'JSON');
  }
  return ContentService.createTextOutput(jsonData).setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Convert CSV to Excel format (SpreadsheetML)
 */
function convertCSVToExcel(csv, dashboardName) {
  const lines = csv.split('\n').filter(line => line.trim().length > 0);
  if (lines.length === 0) return '';
  
  let excelContent = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Title>${dashboardName}</Title>
  <Author>Tableau Extension</Author>
  <Created>${new Date().toISOString()}</Created>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Bottom"/></Style>
  <Style ss:ID="Header"><Font ss:Bold="1"/><Interior ss:Color="#E6F3FF" ss:Pattern="Solid"/></Style>
 </Styles>
 <Worksheet ss:Name="${dashboardName.replace(/[^\w\s]/g, '').substring(0, 31)}">
  <Table>`;
  
  lines.forEach((line, index) => {
    const cells = parseCSVLine(line);
    excelContent += '<Row>\n';
    cells.forEach(cell => {
      // Proper XML entity escaping for Excel
      const cleanCell = cell.replace(/&/g, '&' + 'amp;').replace(/</g, '&' + 'lt;').replace(/>/g, '&' + 'gt;').replace(/"/g, '');
      const styleID = index === 0 ? 'Header' : 'Default';
      const dataType = (index > 0 && !isNaN(cleanCell) && cleanCell.trim() !== '') ? 'Number' : 'String';
      excelContent += `<Cell ss:StyleID="${styleID}"><Data ss:Type="${dataType}">${cleanCell}</Data></Cell>\n`;
    });
    excelContent += '</Row>\n';
  });
  
  excelContent += `</Table></Worksheet></Workbook>`;
  return excelContent;
}

/**
 * Convert CSV to PDF-compatible HTML
 */
function convertCSVToPDF(csv, dashboardName) {
  const lines = csv.split('\n').filter(line => line.trim().length > 0);
  if (lines.length === 0) return '';

  let html = `<!DOCTYPE html><html><head><style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; }
    th { background-color: #f0f8ff; }
  </style></head><body><h1>${dashboardName}</h1><table>`;

  lines.forEach((line, index) => {
    const cells = parseCSVLine(line);
    const tag = index === 0 ? 'th' : 'td';
    html += '<tr>';
    cells.forEach(cell => {
      html += `<${tag}>${cell.replace(/&/g, '&' + 'amp;').replace(/</g, '&' + 'lt;').replace(/>/g, '&' + 'gt;')}</${tag}>`;
    });
    html += '</tr>';
  });

  html += '</table></body></html>';
  return html;
}

/**
 * Convert CSV to JSON
 */
function convertCSVToJSON(csv) {
  const lines = csv.split('\n').filter(line => line.trim().length > 0);
  if (lines.length < 2) return '[]';
  
  const headers = parseCSVLine(lines[0]);
  const rows = lines.slice(1).map(line => {
    const row = {};
    const cells = parseCSVLine(line);
    headers.forEach((header, index) => {
      row[header] = cells[index] || '';
    });
    return row;
  });
  
  return JSON.stringify(rows, null, 2);
}

/**
 * Send email notification
 */
function sendExportNotification(user, dashboard, rowCount, trackerCode, format) {
  if (!SEND_EMAIL_NOTIFICATIONS) return;
  const subject = `📊 Tableau ${format} Export: ${dashboard} [${trackerCode}]`;
  const body = `Export Details:
• Dashboard: ${dashboard}
• User: ${user}
• Format: ${format}
• Timestamp: ${new Date().toLocaleString('en-GB', { timeZone: 'Europe/Brussels' })}
• Rows: ${rowCount}
• Tracker: ${trackerCode}`;
  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
}

/**
 * Generate unique tracker ID
 */
function generateUniqueId() {
  return (Date.now().toString(36) + '-' + Math.random().toString(36).substr(2, 9)).toUpperCase();
}

/**
 * Convert CSV to HTML table for Excel import with formatting
 */
function convertCSVToHTMLTable(csv, dashboardName) {
  const lines = csv.split('\n').filter(line => line.trim().length > 0);
  if (lines.length === 0) return '';

  let html = `<html>
<head>
<meta charset="utf-8">
<style>
table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
th { background-color: #4472C4; color: white; font-weight: bold; border: 1px solid #2F5597; padding: 8px; text-align: left; }
td { border: 1px solid #D0D7DE; padding: 6px; }
tr:nth-child(even) { background-color: #F8F9FA; }
tr:nth-child(odd) { background-color: #FFFFFF; }
</style>
</head>
<body>
<table>`;

  lines.forEach((line, index) => {
    const cells = parseCSVLine(line);
    const tag = index === 0 ? 'th' : 'td';
    html += '<tr>';
    cells.forEach(cell => {
      const cleanCell = cell.replace(/&/g, '&' + 'amp;').replace(/</g, '&' + 'lt;').replace(/>/g, '&' + 'gt;');
      html += `<${tag}>${cleanCell}</${tag}>`;
    });
    html += '</tr>';
  });

  html += '</table></body></html>';
  return html;
}

/**
 * Parse a single line of CSV
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current);
  return result;
}
