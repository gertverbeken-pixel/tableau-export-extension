# Tableau Multi-Format Export Extension

A powerful Tableau dashboard extension that enables intelligent data export in multiple formats (CSV, Excel, PDF, JSON) with automatic filter detection and comprehensive user tracking.

## Features

### 🚀 Multi-Format Export
- **CSV**: Standard comma-separated values format
- **Excel (.xlsx)**: Formatted spreadsheet with styled headers and auto-fitted columns
- **PDF**: Professional table layout with pagination support
- **JSON**: Structured data format with metadata

### 🎯 Smart Filter Detection
- Automatically detects and applies dashboard filters
- Supports multiple filter types and combinations
- Provides readable filter summaries for tracking

### 📊 User Tracking & Analytics
- Tracks user activity and export patterns
- Records dashboard name, timestamp, and applied filters
- Integrates with Google Apps Script for centralized logging

### 🔧 Intelligent Worksheet Selection
- Automatically processes all worksheets in a dashboard
- Handles complex dashboard structures
- Includes comprehensive error handling

## Installation

### Prerequisites
- Node.js 16.0.0 or higher
- Tableau Server or Tableau Online access
- Google Apps Script setup (for current backend)

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/gertverbeken-pixel/tableau-export-extension.git
   cd tableau-export-extension
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Configure environment variables** (for Node.js backend)
   ```bash
   # Create .env file
   TABLEAU_SERVER=https://your-tableau-server.com
   PAT_NAME=your-personal-access-token-name
   PAT_SECRET=your-personal-access-token-secret
   SITE_CONTENT_URL=your-site-content-url
   ```

4. **Deploy to Tableau**
   - Upload `manifest.trex` to your Tableau Server
   - Host `index.html` on a web server (currently using GitHub Pages)

## Usage

### In Tableau Desktop/Server

1. **Add Extension to Dashboard**
   - Drag "Extension" object to your dashboard
   - Choose the manifest file or use the hosted URL
   - Configure permissions if prompted

2. **Select Export Format**
   - Choose from CSV, Excel, PDF, or JSON formats
   - Each format provides optimized output for different use cases

3. **Export Data**
   - Click "Export & Track" button
   - Extension automatically detects applied filters
   - File downloads in selected format with tracking data

### Format-Specific Features

#### Excel Export
- Professional formatting with styled headers
- Auto-fitted column widths
- Suitable for further analysis and sharing

#### PDF Export
- Clean table layout with proper spacing
- Pagination for large datasets
- Ideal for reports and presentations

#### JSON Export
- Structured data with metadata
- Includes export timestamp and row count
- Perfect for API integrations and data processing

## Backend Options

### Current: Google Apps Script
- Handles format conversion and tracking
- No server maintenance required
- Easy to modify and extend

### Alternative: Node.js/Serverless
- Complete `export.js` implementation included
- Supports all export formats with dedicated libraries
- Can be deployed to AWS Lambda, Vercel, or similar platforms

## API Reference

### Query Parameters

The extension sends the following parameters to the backend:

```javascript
{
  searchName: string,      // Dashboard name
  searchType: string,      // Always "dashboard"
  user: string,           // Username or email
  dashboard: string,      // Encoded dashboard name
  maxRows: string,        // Maximum rows to export
  includeAllColumns: boolean,
  universalMode: boolean,
  format: string,         // "csv" | "excel" | "pdf" | "json"
  filters?: string,       // JSON string of applied filters
  filterSummary?: string  // Human-readable filter description
}
```

### Response Formats

- **CSV**: `text/csv` with proper headers
- **Excel**: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- **PDF**: `application/pdf`
- **JSON**: `application/json`

## Development

### Local Development
```bash
npm run dev
```

### Building for Production
```bash
npm run build
```

### Deployment
```bash
npm run deploy
```

## Configuration

### Manifest File
The `manifest.trex` file defines:
- Extension metadata and version
- Required permissions
- Source location URL
- Localization resources

### Frontend Customization
Modify `index.html` to:
- Add new export formats
- Customize UI styling
- Implement additional tracking features

### Backend Extensions
The Node.js backend (`export.js`) can be extended to:
- Add new export formats
- Implement custom data transformations
- Integrate with additional tracking systems

## Dependencies

### Frontend
- Tableau Extensions API
- Modern browser with ES6+ support

### Backend (Node.js)
- `express`: Web framework
- `axios`: HTTP client for Tableau REST API
- `exceljs`: Excel file generation
- `pdfkit`: PDF document creation
- `serverless-http`: Serverless deployment support

## Security Considerations

- Extension requires "full data" permission
- All data transmission uses HTTPS
- User authentication handled by Tableau
- Backend can implement additional security measures

## Troubleshooting

### Common Issues

1. **Extension fails to initialize**
   - Verify Tableau Extensions API is loaded
   - Check browser console for JavaScript errors
   - Ensure proper permissions are granted

2. **Export fails**
   - Verify backend URL is accessible
   - Check network connectivity
   - Review server logs for errors

3. **Format conversion errors**
   - Ensure all required dependencies are installed
   - Check data format and encoding
   - Verify sufficient memory for large datasets

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For issues and questions:
- Email: gert.verbeken@easyfairs.com
- GitHub Issues: [Create an issue](https://github.com/gertverbeken-pixel/tableau-export-extension/issues)

## Changelog

### v1.2.0
- ✨ Added multi-format export support (Excel, PDF, JSON)
- 🎨 Updated UI with format selection dropdown
- 🔧 Enhanced backend with format conversion functions
- 📚 Comprehensive documentation updates

### v1.1.0
- 🚀 Initial CSV export functionality
- 📊 User tracking implementation
- 🎯 Smart filter detection
- 🔌 Google Apps Script integration
