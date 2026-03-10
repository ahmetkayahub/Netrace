import streamlit as st
import os
import pandas as pd
from utils.pdf_processor import (
    extract_p_codes_from_pdf,
    extract_p_codes_from_file_path,
    highlight_ana_pdf,
    highlight_uploaded_pdf,
    extract_p_codes_with_details
)
from utils.importer import (
    import_new_pdfs_to_master,
    get_import_statistics,
    scan_data_folders
)

# Try to import SharePoint connector (optional)
try:
    from utils.sharepoint_connector import SharePointConnector
    SHAREPOINT_AVAILABLE = True
except ImportError:
    SHAREPOINT_AVAILABLE = False

# Page configuration
st.set_page_config(
    page_title="Netrace Order Verification Center",
    page_icon="📄",
    layout="centered"
)

# Main title and description
st.title("Netrace Order Verification Center")
st.markdown("""
### Service and P-Code Analysis

To maintain our operational efficiency, please upload your Purchase Order (PO) document. Our system automatically compares it with the Master list and provides an instant report on service status:

🟢 **Matched**: Verified services found within the PO.

🔴 **Missing**: Codes present in the Master list but not found in the uploaded document.

---
""")

# Master PDF file path
MASTER_PDF_PATH = "data/ana_pdf.pdf"

# Check if Master PDF exists
if not os.path.exists(MASTER_PDF_PATH):
    st.error(f"⚠️ Master PDF file not found: `{MASTER_PDF_PATH}`")
    st.info("Please place the Master PDF file at `data/ana_pdf.pdf`")
    st.stop()

# File uploader
uploaded_file = st.file_uploader(
    "",
    type=['pdf'],
    help="Only PDF files are accepted"
)

# Start processing when file is uploaded
if uploaded_file is not None:
    st.success(f"✅ File uploaded: **{uploaded_file.name}**")
    
    # Process button
    if st.button("🔍 Start Comparison", type="primary"):
        try:
            with st.spinner("⏳ Processing PDF files, please wait..."):
                # 1. Extract P codes from uploaded PDF
                uploaded_pdf_codes = extract_p_codes_from_pdf(uploaded_file)
                
                # 2. Extract P codes from Master PDF (first 6 pages)
                master_pdf_codes = extract_p_codes_from_file_path(MASTER_PDF_PATH, max_pages=6)
                
                # 3. Highlight Master PDF
                highlighted_pdf = highlight_ana_pdf(MASTER_PDF_PATH, uploaded_pdf_codes)
            
            # Statistics
            st.subheader("📊 Comparison Results")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(
                    "Codes in Uploaded PDF",
                    len(uploaded_pdf_codes)
                )
            
            with col2:
                matching_codes = len(uploaded_pdf_codes.intersection(master_pdf_codes))
                st.metric(
                    "Matching Codes (Green)",
                    matching_codes
                )
            
            with col3:
                non_matching_codes = len(master_pdf_codes - uploaded_pdf_codes)
                st.metric(
                    "Other Codes in Master PDF (Red)",
                    non_matching_codes
                )
            
            # Download buttons
            st.subheader("💾 Download Processed PDFs")
            
            col_dl1, col_dl2 = st.columns(2)
            
            with col_dl1:
                st.download_button(
                    label="📥 Download Master PDF (Highlighted)",
                    data=highlighted_pdf,
                    file_name="master_pdf_highlighted.pdf",
                    mime="application/pdf",
                    type="primary"
                )
            
            with col_dl2:
                # Developer feature: Highlight and download uploaded PDF
                try:
                    with st.spinner("🎨 Processing your uploaded PDF..."):
                        highlighted_uploaded = highlight_uploaded_pdf(uploaded_file, master_pdf_codes)
                    
                    st.download_button(
                        label="🔧 Developer: Download Your PDF (Highlighted)",
                        data=highlighted_uploaded,
                        file_name="uploaded_pdf_highlighted.pdf",
                        mime="application/pdf",
                        type="secondary",
                        help="Download your PDF with matching codes highlighted in green"
                    )
                except Exception as e:
                    st.error(f"❌ Error highlighting uploaded PDF: {str(e)}")
            
            # Code details (optional, in expander)
            with st.expander("🔍 Detailed Code List"):
                st.write("**P Codes in Uploaded PDF:**")
                st.code(", ".join(sorted(uploaded_pdf_codes)) if uploaded_pdf_codes else "No codes found")
                
                st.write("**P Codes in Master PDF (First 6 Pages):**")
                st.code(", ".join(sorted(master_pdf_codes)) if master_pdf_codes else "No codes found")
                
                st.write("**Matching Codes (Green):**")
                matching = uploaded_pdf_codes.intersection(master_pdf_codes)
                st.code(", ".join(sorted(matching)) if matching else "No matches")
            
            # Developer feature: Extract P codes with order details
            with st.expander("🔧 Developer: Extract Order Details"):
                st.info("Extract P codes with Order Qty, Price per Unit, and Net Value EUR from uploaded PDF")
                
                if st.button("📊 Extract Details", help="Extract detailed order information"):
                    try:
                        with st.spinner("🔍 Extracting order details..."):
                            details = extract_p_codes_with_details(uploaded_file)
                        
                        if details:
                            st.success(f"✅ Extracted details for {len(details)} P codes")
                            
                            # Convert to DataFrame for better display
                            df = pd.DataFrame(details)
                            st.dataframe(df, use_container_width=True)
                            
                            # Download as CSV
                            csv = df.to_csv(index=False)
                            st.download_button(
                                label="📥 Download as CSV",
                                data=csv,
                                file_name="p_code_details.csv",
                                mime="text/csv"
                            )
                        else:
                            st.warning("⚠️ No P codes with details found in the uploaded PDF")
                    
                    except ValueError as e:
                        st.error(f"❌ Error: {str(e)}")
                    except Exception as e:
                        st.error(f"❌ An unexpected error occurred: {str(e)}")
                        st.exception(e)
        
        except Exception as e:
            st.error(f"❌ Processing failed: {str(e)}")
            st.exception(e)

# ============================================================================
# NEW SECTION: JSON Export for Auto-Import
# ============================================================================

st.markdown("---")
st.header("📤 Export to JSON (Auto-Import)")

st.markdown("""
Extract order details from uploaded PDF and export as JSON for automatic processing.
Upload the JSON file to SharePoint, and the system will automatically import it to Excel.
""")

with st.expander("🔧 Extract Order Details & Export JSON"):
    if uploaded_file is not None:
        if st.button("📊 Extract & Export JSON", use_container_width=True):
            with st.spinner("Extracting data from PDF..."):
                try:
                    import json
                    from datetime import datetime, timezone
                    from utils.po_parser import parse_po_lines_from_pdf
                    
                    # PDF'i byte olarak oku
                    uploaded_file.seek(0)
                    pdf_bytes = uploaded_file.read()
                    
                    # P-code'ları ve detayları çıkar
                    result = parse_po_lines_from_pdf(pdf_bytes)
                    
                    if result and result.get('lines'):
                        # JSON formatında hazırla
                        export_data = {
                            "po_number": result.get('po_number', 'Unknown'),
                            "imported_at": datetime.now(timezone.utc).isoformat(),
                            "source_file": uploaded_file.name,
                            "lines": result['lines']
                        }
                        
                        # JSON string oluştur
                        json_str = json.dumps(export_data, indent=2, ensure_ascii=False)
                        
                        # Sonuçları göster
                        st.success(f"✅ {len(result['lines'])} lines extracted!")
                        
                        # DataFrame olarak göster
                        df = pd.DataFrame(result['lines'])
                        st.dataframe(df, use_container_width=True)
                        
                        # İndirme butonları
                        col_json, col_csv = st.columns(2)
                        
                        with col_json:
                            # JSON indirme butonu
                            json_filename = f"po_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                            st.download_button(
                                label="📥 Download JSON",
                                data=json_str,
                                file_name=json_filename,
                                mime="application/json",
                                use_container_width=True
                            )
                            
                            # Kullanıcıya talimat
                            st.info(f"""
                            **📋 How to Use:**
                            1. Download the JSON file
                            2. Create a folder in SharePoint (e.g., HH1309)
                            3. Upload this JSON file to that folder
                            4. System will automatically import to Excel every 15 minutes!
                            
                            **Filename:** `{json_filename}`
                            """)
                        
                        with col_csv:
                            # CSV indirme (opsiyonel)
                            csv_str = df.to_csv(index=False)
                            st.download_button(
                                label="📥 Download CSV",
                                data=csv_str,
                                file_name=f"po_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
                    else:
                        st.warning("⚠️ No P-codes found in the PDF")
                        
                except Exception as e:
                    st.error(f"❌ Error extracting data: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
    else:
        st.info("👆 Please upload a PDF file first")


# ============================================================================
# NEW SECTION: Import PO PDFs to Master Excel
# ============================================================================

st.markdown("---")
st.header("📥 Import PO PDFs to Master Excel")

st.markdown("""
This feature scans local `data/` folders organized by SahaID and imports PO line items into `Master.xlsx`.

**Folder Structure:**
```
data/
├── HH1309/
│   ├── PO1.pdf
│   └── PO2.pdf
├── HH1310/
│   └── PO3.pdf
└── Master.xlsx
```
""")

col_import1, col_import2 = st.columns(2)

with col_import1:
    master_excel_path = st.text_input(
        "Master Excel Path",
        value="data/Master.xlsx",
        help="Path to Master.xlsx file"
    )

with col_import2:
    data_base_dir = st.text_input(
        "Data Base Directory",
        value="data",
        help="Base directory containing SahaID folders"
    )

# Show current statistics
if st.button("📊 Show Import Statistics"):
    try:
        stats = get_import_statistics("data/import_state.json")
        
        st.subheader("Import Statistics")
        col_stat1, col_stat2 = st.columns(2)
        
        with col_stat1:
            st.metric("Total Files Imported", stats['total_files_imported'])
        
        with col_stat2:
            st.metric("Total Lines Imported", stats['total_lines_imported'])
        
        if stats['files_by_saha']:
            st.write("**Files by SahaID:**")
            for saha_id, count in stats['files_by_saha'].items():
                st.write(f"- {saha_id}: {count} files")
    
    except Exception as e:
        st.info("No import history yet")

# Scan and Import button
if st.button("🔍 Scan & Import", type="primary"):
    try:
        with st.spinner("⏳ Scanning folders and importing PDFs..."):
            # Scan folders first
            pdf_files = scan_data_folders(data_base_dir)
            
            if not pdf_files:
                st.warning(f"⚠️ No PDF files found in `{data_base_dir}/` subfolders")
                st.info("Please create SahaID folders (e.g., data/HH1309/) and add PDF files")
            else:
                st.info(f"Found {len(pdf_files)} PDF file(s) to process...")
                
                # Import to Master
                result = import_new_pdfs_to_master(
                    base_dir=data_base_dir,
                    master_path=master_excel_path,
                    state_path="data/import_state.json"
                )
                
                # Display results
                st.success("✅ Import process completed!")
                
                st.subheader("📊 Import Results")
                
                col_res1, col_res2, col_res3, col_res4 = st.columns(4)
                
                with col_res1:
                    st.metric("PDFs Found", result['total_pdfs_found'])
                
                with col_res2:
                    st.metric("New PDFs Processed", result['new_pdfs_processed'])
                
                with col_res3:
                    st.metric("Lines Added", result['total_lines_added'])
                
                with col_res4:
                    st.metric("PDFs Skipped", result['skipped_pdfs'])
                
                # Show errors if any
                if result['errors']:
                    st.warning(f"⚠️ {len(result['errors'])} file(s) had errors:")
                    
                    error_df = pd.DataFrame(result['errors'])
                    st.dataframe(error_df, use_container_width=True)
                else:
                    st.success("🎉 All files processed successfully!")
                
                # Show success details
                if result['new_pdfs_processed'] > 0:
                    st.info(f"✅ Successfully added {result['total_lines_added']} line(s) to `{master_excel_path}`")
    
    except Exception as e:
        st.error(f"❌ Import error: {str(e)}")
        st.exception(e)


# ============================================================================
# SharePoint Integration (Optional - for Streamlit Cloud)
# ============================================================================

if SHAREPOINT_AVAILABLE:
    st.markdown("---")
    st.header("☁️ SharePoint Integration")
    
    # Check if secrets are configured
    sharepoint_configured = False
    if hasattr(st, 'secrets') and 'sharepoint' in st.secrets:
        sharepoint_configured = st.secrets['sharepoint'].get('enabled', False)
    
    if sharepoint_configured:
        st.success("✅ SharePoint integration enabled")
        
        col_sp1, col_sp2 = st.columns(2)
        
        with col_sp1:
            if st.button("📥 Download PDFs from SharePoint"):
                try:
                    with st.spinner("🌐 Connecting to SharePoint..."):
                        sp = SharePointConnector(
                            client_id=st.secrets['sharepoint']['client_id'],
                            client_secret=st.secrets['sharepoint']['client_secret'],
                            tenant_id=st.secrets['sharepoint']['tenant_id'],
                            site_id=st.secrets['sharepoint']['site_id']
                        )
                        
                        base_folder = st.secrets['sharepoint'].get('base_folder', '/Documents')
                        
                        sp.authenticate()
                        st.success("✅ Connected to SharePoint")
                        
                        downloaded = sp.download_saha_pdfs(base_folder, 'data')
                        
                        if downloaded:
                            st.success(f"✅ Downloaded {len(downloaded)} PDF file(s)")
                            df = pd.DataFrame(downloaded, columns=['SahaID', 'File Path'])
                            st.dataframe(df, use_container_width=True)
                        else:
                            st.info("ℹ️ No new PDFs found in SharePoint")
                
                except Exception as e:
                    st.error(f"❌ SharePoint error: {str(e)}")
        
        with col_sp2:
            if st.button("📤 Upload Master.xlsx to SharePoint"):
                try:
                    with st.spinner("🌐 Uploading to SharePoint..."):
                        sp = SharePointConnector(
                            client_id=st.secrets['sharepoint']['client_id'],
                            client_secret=st.secrets['sharepoint']['client_secret'],
                            tenant_id=st.secrets['sharepoint']['tenant_id'],
                            site_id=st.secrets['sharepoint']['site_id']
                        )
                        
                        master_path_sp = st.secrets['sharepoint'].get('master_path', '/Documents/Master.xlsx')
                        
                        sp.authenticate()
                        
                        # Read local file
                        with open('data/Master.xlsx', 'rb') as f:
                            file_content = f.read()
                        
                        # Upload
                        result = sp.upload_file(master_path_sp, file_content, overwrite=True)
                        
                        st.success(f"✅ Uploaded to: {master_path_sp}")
                
                except Exception as e:
                    st.error(f"❌ Upload error: {str(e)}")
    else:
        st.info("ℹ️ SharePoint integration not configured")
        st.markdown("""
        To enable SharePoint integration:
        1. Configure secrets in Streamlit Cloud settings
        2. See `.streamlit/secrets.toml.example` for required values
        """)



