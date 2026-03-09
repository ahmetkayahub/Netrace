# utils/importer.py

import os
import json
import hashlib
from pathlib import Path
from typing import List, Tuple, Dict
from .po_parser import parse_po_lines_from_pdf, validate_po_data
from .excel_writer import append_saha_kayitlari_rows


def calculate_file_hash(file_path: str) -> str:
    """
    Calculates SHA256 hash of a file.
    
    Args:
        file_path: Path to file
    
    Returns:
        str: SHA256 hash
    """
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()


def load_import_state(state_path: str) -> Dict:
    """
    Loads import state from JSON file.
    
    Args:
        state_path: Path to import_state.json
    
    Returns:
        dict: Import state {file_path: {hash, mtime, size}}
    """
    if os.path.exists(state_path):
        try:
            with open(state_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_import_state(state_path: str, state: Dict) -> None:
    """
    Saves import state to JSON file.
    
    Args:
        state_path: Path to import_state.json
        state: Import state dictionary
    """
    # Ensure directory exists
    Path(state_path).parent.mkdir(parents=True, exist_ok=True)
    
    with open(state_path, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def scan_data_folders(base_dir: str = "data") -> List[Tuple[str, str]]:
    """
    Scans data folders for PDF files organized by SahaID.
    
    Args:
        base_dir: Base directory to scan (default: "data")
    
    Returns:
        List[Tuple[str, str]]: List of (saha_id, pdf_path) tuples
    
    Example:
        data/HH1309/PO1.pdf -> ("HH1309", "data/HH1309/PO1.pdf")
    """
    results = []
    base_path = Path(base_dir)
    
    if not base_path.exists():
        return results
    
    # Scan subdirectories
    for saha_dir in base_path.iterdir():
        if saha_dir.is_dir() and not saha_dir.name.startswith('.'):
            saha_id = saha_dir.name
            
            # Find all PDF files in this directory
            for pdf_file in saha_dir.glob("*.pdf"):
                if pdf_file.is_file():
                    results.append((saha_id, str(pdf_file)))
    
    return results


def import_new_pdfs_to_master(base_dir: str = "data", 
                              master_path: str = "data/Master.xlsx",
                              state_path: str = "data/import_state.json") -> Dict:
    """
    Imports new PDFs from data folders to Master Excel.
    
    Args:
        base_dir: Base directory to scan
        master_path: Path to Master.xlsx
        state_path: Path to import_state.json
    
    Returns:
        dict: {
            'total_pdfs_found': int,
            'new_pdfs_processed': int,
            'total_lines_added': int,
            'skipped_pdfs': int,
            'errors': List[dict]  # [{file, error}]
        }
    """
    # Load import state
    import_state = load_import_state(state_path)
    
    # Scan folders
    pdf_files = scan_data_folders(base_dir)
    
    result = {
        'total_pdfs_found': len(pdf_files),
        'new_pdfs_processed': 0,
        'total_lines_added': 0,
        'skipped_pdfs': 0,
        'errors': []
    }
    
    # Process each PDF
    for saha_id, pdf_path in pdf_files:
        try:
            # Calculate file hash
            file_hash = calculate_file_hash(pdf_path)
            file_stat = os.stat(pdf_path)
            
            # Check if already imported
            if pdf_path in import_state:
                stored_hash = import_state[pdf_path].get('hash')
                if stored_hash == file_hash:
                    # File already imported and unchanged
                    result['skipped_pdfs'] += 1
                    continue
            
            # Read PDF
            with open(pdf_path, 'rb') as f:
                pdf_bytes = f.read()
            
            # Parse PO
            po_data = parse_po_lines_from_pdf(pdf_bytes)
            
            # Validate
            if not validate_po_data(po_data):
                result['errors'].append({
                    'file': pdf_path,
                    'error': 'No valid lines found in PDF'
                })
                continue
            
            # Prepare rows for Excel
            rows_to_add = []
            po_number = po_data.get('po_number', '')
            source_filename = os.path.basename(pdf_path)
            
            for line in po_data['lines']:
                if line.get('HizmetKodu'):  # Only add lines with HizmetKodu
                    row = {
                        'SahaID': saha_id,
                        'PONumber': po_number,
                        'HizmetKodu': line.get('HizmetKodu', ''),
                        'OrderQty': line.get('OrderQty', ''),
                        'Price': line.get('Price', ''),
                        'NetValue': line.get('NetValue', ''),
                        'SourceFileName': source_filename
                    }
                    rows_to_add.append(row)
            
            if rows_to_add:
                # Append to Master Excel
                lines_added = append_saha_kayitlari_rows(master_path, rows_to_add)
                result['total_lines_added'] += lines_added
                result['new_pdfs_processed'] += 1
                
                # Update import state
                import_state[pdf_path] = {
                    'hash': file_hash,
                    'mtime': file_stat.st_mtime,
                    'size': file_stat.st_size,
                    'saha_id': saha_id,
                    'lines_imported': len(rows_to_add)
                }
            else:
                result['errors'].append({
                    'file': pdf_path,
                    'error': 'No valid HizmetKodu lines found'
                })
        
        except Exception as e:
            result['errors'].append({
                'file': pdf_path,
                'error': str(e)
            })
    
    # Save import state
    save_import_state(state_path, import_state)
    
    return result


def get_import_statistics(state_path: str = "data/import_state.json") -> Dict:
    """
    Gets statistics from import state.
    
    Args:
        state_path: Path to import_state.json
    
    Returns:
        dict: Statistics about imported files
    """
    import_state = load_import_state(state_path)
    
    total_files = len(import_state)
    total_lines = sum(item.get('lines_imported', 0) for item in import_state.values())
    
    saha_counts = {}
    for item in import_state.values():
        saha_id = item.get('saha_id', 'Unknown')
        saha_counts[saha_id] = saha_counts.get(saha_id, 0) + 1
    
    return {
        'total_files_imported': total_files,
        'total_lines_imported': total_lines,
        'files_by_saha': saha_counts
    }
