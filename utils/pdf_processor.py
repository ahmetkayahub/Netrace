# utils/pdf_processor.py

import fitz  # PyMuPDF
import re
from typing import Set, Optional
import io


def extract_p_codes_from_pdf(pdf_file, max_pages: Optional[int] = None) -> Set[str]:
    """
    Extracts P codes from a PDF file.
    
    Args:
        pdf_file: PDF file (file-like object or bytes)
        max_pages: Maximum number of pages to process (None = all pages)
    
    Returns:
        Set[str]: Unique list of P codes found (e.g., {'P123456', 'P678901'})
    
    Raises:
        ValueError: If PDF cannot be read
    """
    p_codes = set()
    
    try:
        # Read PDF from memory
        if isinstance(pdf_file, bytes):
            pdf_bytes = pdf_file
        else:
            pdf_bytes = pdf_file.read()
        
        # Open PDF
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Determine page count
        total_pages = len(pdf_document)
        pages_to_process = min(max_pages, total_pages) if max_pages else total_pages
        
        # Regex pattern: P + exactly 6 digits
        pattern = r'P\d{6}'
        
        # Scan each page
        for page_num in range(pages_to_process):
            page = pdf_document[page_num]
            text = page.get_text()
            
            # Find P codes
            matches = re.findall(pattern, text)
            p_codes.update(matches)
        
        pdf_document.close()
        
        return p_codes
    
    except Exception as e:
        raise ValueError(f"Error processing PDF: {str(e)}")


def extract_p_codes_from_file_path(file_path: str, max_pages: Optional[int] = None) -> Set[str]:
    """
    Reads PDF from file path and extracts P codes.
    
    Args:
        file_path: Path to PDF file
        max_pages: Maximum number of pages to process (None = all pages)
    
    Returns:
        Set[str]: Unique list of P codes found
    
    Raises:
        ValueError: If PDF cannot be read
    """
    try:
        with open(file_path, 'rb') as f:
            return extract_p_codes_from_pdf(f, max_pages)
    except FileNotFoundError:
        raise ValueError(f"PDF file not found: {file_path}")
    except Exception as e:
        raise ValueError(f"Error reading PDF file: {str(e)}")


def highlight_ana_pdf(ana_pdf_path: str, uploaded_pdf_codes: Set[str]) -> io.BytesIO:
    """
    Finds P codes in first 6 pages of Master PDF and highlights them by comparison.
    
    Args:
        ana_pdf_path: Path to Master PDF file on server
        uploaded_pdf_codes: P codes from user's uploaded PDF (set)
    
    Returns:
        io.BytesIO: Highlighted PDF in memory
    
    Raises:
        ValueError: If error occurs during PDF processing
    """
    try:
        # Open Master PDF
        pdf_document = fitz.open(ana_pdf_path)
        
        # Process only first 6 pages
        max_pages = min(6, len(pdf_document))
        
        # Regex pattern: P + exactly 6 digits
        pattern = re.compile(r'P\d{6}')
        
        # Color definitions (RGB format, values 0-1)
        GREEN_HIGHLIGHT = [0.0, 1.0, 0.0]   # Green
        RED_HIGHLIGHT = [1.0, 0.0, 0.0]     # Red
        
        # Process each page
        for page_num in range(max_pages):
            page = pdf_document[page_num]
            
            # Get text with position information
            text_instances = page.get_text("words")  # Returns list of (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            
            # Search for P codes in text instances
            for inst in text_instances:
                word = inst[4]  # The text content
                
                # Check if this word matches P code pattern
                if pattern.fullmatch(word):
                    # Get bounding box coordinates
                    rect = fitz.Rect(inst[0], inst[1], inst[2], inst[3])
                    
                    # Select highlight color
                    if word in uploaded_pdf_codes:
                        highlight_color = GREEN_HIGHLIGHT
                    else:
                        highlight_color = RED_HIGHLIGHT
                    
                    # Add highlight
                    highlight = page.add_highlight_annot(rect)
                    highlight.set_colors(stroke=highlight_color)
                    highlight.update()
        
        # Save PDF to memory
        pdf_bytes = io.BytesIO()
        pdf_document.save(pdf_bytes)
        pdf_document.close()
        
        # Reset BytesIO cursor to beginning
        pdf_bytes.seek(0)
        
        return pdf_bytes
    
    except FileNotFoundError:
        raise ValueError(f"Master PDF file not found: {ana_pdf_path}")
    except Exception as e:
        raise ValueError(f"Error highlighting Master PDF: {str(e)}")


def highlight_uploaded_pdf(uploaded_pdf_file, master_pdf_codes: Set[str]) -> io.BytesIO:
    """
    Highlights P codes in uploaded PDF that match with Master PDF codes.
    
    Args:
        uploaded_pdf_file: Uploaded PDF file (file-like object or bytes)
        master_pdf_codes: P codes from Master PDF (set)
    
    Returns:
        io.BytesIO: Highlighted PDF in memory
    
    Raises:
        ValueError: If error occurs during PDF processing
    """
    try:
        # Read PDF from memory
        if isinstance(uploaded_pdf_file, bytes):
            pdf_bytes = uploaded_pdf_file
        else:
            uploaded_pdf_file.seek(0)  # Reset file pointer
            pdf_bytes = uploaded_pdf_file.read()
        
        # Open PDF
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Regex pattern: P + exactly 6 digits
        pattern = re.compile(r'P\d{6}')
        
        # Color definition (RGB format, values 0-1)
        GREEN_HIGHLIGHT = [0.0, 1.0, 0.0]   # Green
        
        # Process each page
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Get text with position information
            text_instances = page.get_text("words")  # Returns list of (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            
            # Search for P codes in text instances
            for inst in text_instances:
                word = inst[4]  # The text content
                
                # Check if this word matches P code pattern and is in master PDF
                if pattern.fullmatch(word) and word in master_pdf_codes:
                    # Get bounding box coordinates
                    rect = fitz.Rect(inst[0], inst[1], inst[2], inst[3])
                    
                    # Add green highlight
                    highlight = page.add_highlight_annot(rect)
                    highlight.set_colors(stroke=GREEN_HIGHLIGHT)
                    highlight.update()
        
        # Save PDF to memory
        pdf_bytes = io.BytesIO()
        pdf_document.save(pdf_bytes)
        pdf_document.close()
        
        # Reset BytesIO cursor to beginning
        pdf_bytes.seek(0)
        
        return pdf_bytes
    
    except Exception as e:
        raise ValueError(f"Error highlighting uploaded PDF: {str(e)}")


