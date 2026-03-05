import streamlit as st
import os
from utils.pdf_processor import (
    extract_p_codes_from_pdf,
    extract_p_codes_from_file_path,
    highlight_ana_pdf,
    highlight_uploaded_pdf
)

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
        
        except ValueError as e:
            st.error(f"❌ Error: {str(e)}")
        except Exception as e:
            st.error(f"❌ An unexpected error occurred: {str(e)}")
            st.exception(e)



