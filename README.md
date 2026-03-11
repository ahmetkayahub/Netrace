# Netrace Order Verification Center

A professional PDF P-code comparison and verification system that automatically analyzes Purchase Order documents against a Master list.

## Features

- 🎯 **Automated P-Code Extraction**: Automatically extracts P-codes (format: P123456) from PDF documents
- 🔍 **Smart Comparison**: Compares uploaded PO documents against Master list (first 6 pages)
- 🎨 **Visual Highlighting**: 
  - 🟢 Green: Matched codes (found in both documents)
  - 🔴 Red: Missing codes (in Master but not in uploaded document)
- 📊 **Detailed Analytics**: Real-time statistics and detailed code lists
- 🔧 **Developer Mode**: Additional feature to highlight matched codes in uploaded documents
- 💻 **Modern Web Interface**: Built with Streamlit for a clean, professional experience
- ☁️ **Cloud-Ready**: Memory-based processing, perfect for deployment

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
streamlit run app.py
```

Then open your browser and navigate to `http://localhost:8501`

## Project Structure

```
pdf-highlighter/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # Documentation
├── utils/
│   ├── __init__.py
│   └── pdf_processor.py  # PDF processing and highlighting functions
└── data/
    └── ana_pdf.pdf       # Master PDF file (reference document)
```

## How It Works

1. **Upload**: User uploads their Purchase Order (PO) PDF document
2. **Extract**: System extracts all P-codes from the uploaded document
3. **Compare**: Compares extracted codes with Master PDF (first 6 pages)
4. **Highlight**: 
   - Master PDF: Green for matched codes, Red for missing codes
   - Uploaded PDF (Developer mode): Green for codes found in Master
5. **Download**: Provides highlighted PDFs for download and analysis

## Technologies

- **Python 3.8+**
- **Streamlit** - Modern web interface framework
- **PyMuPDF (fitz)** - Advanced PDF processing and manipulation
- **Regular Expressions** - P-code pattern matching

## Requirements

- Python 3.8 or higher
- Internet connection (for initial setup)
- PDF files with P-codes in format: P followed by 6 digits (e.g., P123456)

## Configuration

Place your Master PDF file at: `data/ana_pdf.pdf`

The system will automatically process the first 6 pages of this document as the reference.

## License

This project is provided as-is for order verification purposes.

## Support

For issues or questions, please refer to the project documentation or contact the development team.

