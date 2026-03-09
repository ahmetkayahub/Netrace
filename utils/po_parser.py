# utils/po_parser.py

import fitz  # PyMuPDF
import re
from typing import Optional, List, Dict


def parse_po_lines_from_pdf(pdf_bytes: bytes) -> Dict:
    """
    Parses PO PDF and extracts PO number and line items.
    
    Args:
        pdf_bytes: PDF file as bytes
    
    Returns:
        dict: {
            "po_number": str | None,
            "lines": list[dict]  # each dict has: HizmetKodu, OrderQty, Price, NetValue
        }
    
    Raises:
        ValueError: If PDF cannot be read or parsed
    """
    try:
        # Open PDF
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Extract PO number (search in first page)
        po_number = None
        if len(pdf_document) > 0:
            first_page_text = pdf_document[0].get_text()
            # Try to find PO number pattern (common patterns: PO #, Purchase Order, etc.)
            po_patterns = [
                r'(?:PO|Purchase Order|P\.O\.|Order)\s*[#:№]?\s*(\d{10})',  # 10 digit PO
                r'(?:PO|Purchase Order|P\.O\.|Order)\s*[#:№]?\s*([A-Z0-9]{8,12})',  # Alphanumeric PO
            ]
            for pattern in po_patterns:
                match = re.search(pattern, first_page_text, re.IGNORECASE)
                if match:
                    po_number = match.group(1)
                    break
        
        # Regex patterns
        hizmet_kodu_pattern = re.compile(r'P\d{6}')  # P + 6 digits
        
        lines = []
        
        # Process all pages
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Get text with position information
            words = page.get_text("words")  # (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            
            # Find HizmetKodu (P codes)
            for i, word in enumerate(words):
                text = word[4]
                
                # Check if this word matches HizmetKodu pattern
                if hizmet_kodu_pattern.fullmatch(text):
                    hizmet_kodu = text
                    y_coord = word[1]  # Y coordinate
                    
                    # Find words on the same line (within 5 pixel tolerance)
                    same_line_words = [
                        w for w in words 
                        if abs(w[1] - y_coord) < 5  # Same line
                    ]
                    
                    # Sort by X position (left to right)
                    same_line_words.sort(key=lambda w: w[0])
                    
                    # Find P code index
                    p_code_index = next(
                        (idx for idx, w in enumerate(same_line_words) if w[4] == hizmet_kodu), 
                        None
                    )
                    
                    if p_code_index is not None:
                        # Get values after P code
                        values_after = [w[4] for w in same_line_words[p_code_index + 1:]]
                        
                        # Extract numeric values
                        numeric_values = []
                        for val in values_after:
                            # Clean and check if it's a number
                            cleaned_val = val.replace(',', '').replace('.', '', val.count('.') - 1)
                            if re.search(r'\d', val):
                                numeric_values.append(val)
                        
                        # Extract order details (best-effort)
                        # Typically: OrderQty, Price, NetValue
                        order_qty = None
                        price = None
                        net_value = None
                        
                        # Try to identify values based on position and format
                        for idx, val in enumerate(numeric_values):
                            # First numeric is likely OrderQty (small integer)
                            if idx == 0 and order_qty is None:
                                # Check if it's a small integer (typical for quantity)
                                try:
                                    qty_val = val.replace(',', '')
                                    if '.' not in qty_val or qty_val.split('.')[1] == '00':
                                        order_qty = val
                                        continue
                                except:
                                    pass
                            
                            # Look for price-like values (decimal numbers)
                            if '.' in val or ',' in val:
                                if price is None:
                                    price = val
                                elif net_value is None:
                                    net_value = val
                        
                        # Create line item
                        line_item = {
                            'HizmetKodu': hizmet_kodu,
                            'OrderQty': order_qty,
                            'Price': price,
                            'NetValue': net_value
                        }
                        
                        lines.append(line_item)
        
        pdf_document.close()
        
        return {
            "po_number": po_number,
            "lines": lines
        }
    
    except Exception as e:
        raise ValueError(f"Error parsing PO PDF: {str(e)}")


def validate_po_data(po_data: Dict) -> bool:
    """
    Validates parsed PO data.
    
    Args:
        po_data: Dictionary returned by parse_po_lines_from_pdf
    
    Returns:
        bool: True if data is valid
    """
    if not isinstance(po_data, dict):
        return False
    
    if 'lines' not in po_data:
        return False
    
    if not isinstance(po_data['lines'], list):
        return False
    
    # At least one line should have HizmetKodu
    has_valid_line = any(
        isinstance(line, dict) and line.get('HizmetKodu') 
        for line in po_data['lines']
    )
    
    return has_valid_line
