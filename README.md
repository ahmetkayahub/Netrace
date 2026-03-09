# Netrace Order Verification Center

A professional PDF P-code comparison and verification system that automatically analyzes Purchase Order documents against a Master list, with **JSON-based auto-import** and **SharePoint integration** (no admin consent required).

## 🌟 Features

### P-Code Verification
- 🎯 **Automated P-Code Extraction**: Automatically extracts P-codes (format: P123456) from PDF documents
- 🔍 **Smart Comparison**: Compares uploaded PO documents against Master list (first 6 pages)
- 🎨 **Visual Highlighting**: 
  - 🟢 Green: Matched codes (found in both documents)
  - 🔴 Red: Missing codes (in Master but not in uploaded document)
- 📊 **Detailed Analytics**: Real-time statistics and detailed code lists
- 🔧 **Developer Mode**: Additional feature to highlight matched codes in uploaded documents

### JSON Export & Auto-Import (NEW! ⚡)
- 📤 **Automatic Data Extraction**: Upload PDF → Extract order details → Download JSON
- 📋 **Structured Data**: JSON format with PO number, line items, quantities, prices
- ☁️ **SharePoint Sync**: Upload JSON to SharePoint → Auto-import every 15 minutes
- 🔐 **No Admin Consent Required**: Uses OneDrive sync instead of SharePoint API
- 💰 **100% FREE**: GitHub Actions free tier (2,000 min/month)

### Excel Import (Classic Method)
- 📥 **Automated PO Import**: Scans local folders organized by SahaID and imports PO line items to Master.xlsx
- 🔄 **Idempotent Processing**: Tracks imported files to prevent duplicates
- 📁 **Folder-Based Organization**: Supports data/HH1309/, data/HH1310/ structure
- 📋 **Detailed Line Extraction**: Extracts HizmetKodu, OrderQty, Price, NetValue from PO PDFs
- 📊 **Import Statistics**: View import history and statistics

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Start the Application

```bash
streamlit run app.py
```

Then open your browser and navigate to `http://localhost:8501`

### P-Code Verification

1. Upload your PO PDF document
2. Click "Start Comparison"
3. Download highlighted PDFs

### JSON Export & Auto-Import

**User Workflow:**
1. Upload PDF to https://netrace.streamlit.app/
2. Click "Extract & Export JSON"
3. Download JSON file
4. Upload to SharePoint (e.g., Documents/HH1309/)
5. Wait 15 minutes → Auto-imported to Master.xlsx!

**Setup:**
See [JSON_IMPORT_SETUP.md](JSON_IMPORT_SETUP.md) for detailed setup instructions.
4. View detailed analytics

### Excel Import

1. **Folder Structure**: Create SahaID folders in `data/` directory:
   ```
   data/
   ├── HH1309/
   │   ├── PO1.pdf
   │   └── PO2.pdf
   ├── HH1310/
   │   └── PO3.pdf
   └── Master.xlsx
   ```

2. **Import Process**:
   - Navigate to "Import PO PDFs to Master Excel" section
   - Verify paths (default: `data/Master.xlsx` and `data/`)
   - Click "Scan & Import"
   - View import results and statistics

3. **Master.xlsx Structure**:
   - **Katalog** sheet: Manually managed catalog
   - **SahaKayitlari** sheet: Auto-populated with imported data
     - Columns: SahaID, PONumber, HizmetKodu, OrderQty, Price, NetValue, SourceFileName, ImportedAt

## Project Structure

```
pdf-highlighter/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # Documentation
├── utils/
│   ├── __init__.py
│   ├── pdf_processor.py  # PDF processing and highlighting
│   ├── po_parser.py      # PO line parsing (NEW)
│   ├── excel_writer.py   # Excel operations (NEW)
│   └── importer.py       # Import orchestration (NEW)
└── data/
    ├── ana_pdf.pdf       # Master PDF file (reference document)
    ├── Master.xlsx       # Master Excel workbook (auto-created)
    ├── import_state.json # Import tracking (auto-created)
    ├── HH1309/          # SahaID folders (create as needed)
    └── HH1310/          # SahaID folders (create as needed)
```

## How It Works

### P-Code Verification Flow
1. **Upload**: User uploads their Purchase Order (PO) PDF document
2. **Extract**: System extracts all P-codes from the uploaded document
3. **Compare**: Compares extracted codes with Master PDF (first 6 pages)
4. **Highlight**: 
   - Master PDF: Green for matched codes, Red for missing codes
   - Uploaded PDF (Developer mode): Green for codes found in Master
5. **Download**: Provides highlighted PDFs for download and analysis

### Excel Import Flow
1. **Scan**: Searches `data/` directory for SahaID folders and PDF files
2. **Parse**: Extracts PO number, HizmetKodu (P123456), OrderQty, Price, NetValue
3. **Track**: Checks `import_state.json` to avoid duplicate imports (using file hash)
4. **Append**: Writes new lines to Master.xlsx > SahaKayitlari sheet
5. **Report**: Shows statistics and errors

## Technologies

- **Python 3.8+**
- **Streamlit** - Modern web interface framework
- **PyMuPDF (fitz)** - Advanced PDF processing and manipulation
- **openpyxl** - Excel file operations
- **pandas** - Data manipulation and display
- **Regular Expressions** - P-code pattern matching

## Requirements

- Python 3.8 or higher
- Internet connection (for initial setup)
- PDF files with P-codes in format: P followed by 6 digits (e.g., P123456)

## 📋 Configuration

### Master PDF
Place your Master PDF file at: `data/ana_pdf.pdf`

The system will automatically process the first 6 pages of this document as the reference.

### Master Excel
The `Master.xlsx` file will be auto-created in `data/` directory with:
- **Katalog** sheet (empty, for manual management)
- **SahaKayitlari** sheet (auto-populated during import)

### Import State
Import tracking is stored in `data/import_state.json` (auto-created).

### SharePoint Integration
See **[SHAREPOINT_SETUP.md](SHAREPOINT_SETUP.md)** for detailed setup instructions.

Quick summary:
1. Create Azure AD app registration
2. Configure GitHub Secrets
3. Enable GitHub Actions workflow
4. (Optional) Deploy to Streamlit Cloud

## 🏗️ Architecture

### Local Mode
```
data/
├── HH1309/*.pdf → Parser → Master.xlsx
└── HH1310/*.pdf → Parser → Master.xlsx
```

### SharePoint + GitHub Actions (Automated)
```
SharePoint Online (HH1309/, HH1310/)
    ↓ (every 15 min - GitHub Actions)
Download PDFs → Process → Update Master.xlsx
    ↓
Upload Master.xlsx → SharePoint Online
```

### Streamlit Cloud (Manual UI)
```
User → Streamlit UI → SharePoint API
    ↓
Download/Upload PDFs
    ↓
Process & Display Results
```

## 📝 Notes

- **SharePoint Integration**: Fully automated with GitHub Actions (100% FREE)
- **PO Parser**: Currently uses best-effort extraction for Qty/Price/NetValue. Will be improved based on actual PO formats.
- **Idempotent**: Same PDF file will not be imported twice (tracked by SHA256 hash)
- **Error Handling**: Application continues on errors and reports them in the UI
- **Security**: Uses OAuth 2.0 with Microsoft Graph API

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is provided as-is for order verification purposes.

## Support

For issues or questions, please refer to the project documentation or contact the development team.


