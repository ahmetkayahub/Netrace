#!/usr/bin/env python3
"""
SharePoint klasörlerindeki JSON dosyalarını okuyup Master.xlsx'e aktarır.
OneDrive Sync ile çalışır - SharePoint API GEREKMİYOR!

Kullanım:
    python3 scripts/auto_import_json.py
    
Klasör Yapısı:
    data/
    ├── HH1309/
    │   ├── po_data_20240301_120000.json
    │   └── po_data_20240302_130000.json
    ├── HH1310/
    │   └── po_data_20240303_140000.json
    └── Master.xlsx
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime, timezone

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from utils.excel_writer import append_saha_kayitlari_rows, ensure_master_workbook


def process_json_files(base_dir: str = "data", master_path: str = "data/Master.xlsx"):
    """
    data/ klasörü altındaki tüm JSON dosyalarını işle ve Master.xlsx'e aktar.
    
    Args:
        base_dir: Tarama yapılacak ana klasör (varsayılan: "data")
        master_path: Master Excel dosyası yolu (varsayılan: "data/Master.xlsx")
    
    Returns:
        dict: İşlem sonucu istatistikleri
    """
    all_rows = []
    processed_files = []
    errors = []
    
    print(f"📂 Scanning directory: {base_dir}")
    
    # Master Excel dosyasının varlığını kontrol et
    ensure_master_workbook(master_path)
    
    # data/ altındaki tüm klasörleri tara
    if not os.path.exists(base_dir):
        print(f"⚠️  Base directory not found: {base_dir}")
        return {
            'total_files': 0,
            'rows_added': 0,
            'errors': [f"Base directory not found: {base_dir}"]
        }
    
    for item in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item)
        
        # Klasör mü kontrol et (Master.xlsx dosyasını atla)
        if os.path.isdir(item_path):
            saha_id = item
            
            # Klasördeki tüm .json dosyalarını bul
            for file_name in os.listdir(item_path):
                if file_name.endswith('.json'):
                    file_path = os.path.join(item_path, file_name)
                    
                    try:
                        print(f"📄 Processing: {file_path}")
                        
                        # JSON dosyasını oku
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        
                        # Verileri Excel formatına çevir
                        po_number = data.get('po_number', 'Unknown')
                        lines = data.get('lines', [])
                        
                        if not lines:
                            print(f"  ⚠️  No lines found in JSON")
                            continue
                        
                        for line in lines:
                            all_rows.append({
                                'SahaID': saha_id,
                                'PONumber': po_number,
                                'HizmetKodu': line.get('HizmetKodu', ''),
                                'OrderQty': line.get('OrderQty', ''),
                                'Price': line.get('Price', ''),
                                'NetValue': line.get('NetValue', ''),
                                'SourceFileName': file_name,
                                'ImportedAt': data.get('imported_at', datetime.now(timezone.utc).isoformat())
                            })
                        
                        processed_files.append(file_path)
                        print(f"  ✅ {len(lines)} lines extracted")
                        
                    except json.JSONDecodeError as e:
                        error_msg = f"{file_path}: Invalid JSON format - {str(e)}"
                        errors.append(error_msg)
                        print(f"  ❌ Error: {error_msg}")
                    except Exception as e:
                        error_msg = f"{file_path}: {str(e)}"
                        errors.append(error_msg)
                        print(f"  ❌ Error: {error_msg}")
    
    # Excel'e yaz
    if all_rows:
        try:
            print(f"\n📝 Writing {len(all_rows)} rows to {master_path}...")
            append_saha_kayitlari_rows(master_path, all_rows)
            print(f"✅ Successfully added {len(all_rows)} rows to {master_path}")
        except Exception as e:
            error_msg = f"Excel write error: {str(e)}"
            print(f"\n❌ {error_msg}")
            errors.append(error_msg)
    else:
        print("\n⚠️  No JSON files found or no valid data to import")
    
    # Özet rapor
    print(f"\n{'='*60}")
    print(f"📊 IMPORT SUMMARY")
    print(f"{'='*60}")
    print(f"  Files processed:    {len(processed_files)}")
    print(f"  Rows added:         {len(all_rows)}")
    print(f"  Errors encountered: {len(errors)}")
    print(f"{'='*60}")
    
    if errors:
        print("\n❌ ERRORS:")
        for i, err in enumerate(errors, 1):
            print(f"  {i}. {err}")
    
    return {
        'total_files': len(processed_files),
        'rows_added': len(all_rows),
        'errors': errors
    }


if __name__ == "__main__":
    # Argüman desteği
    import argparse
    
    parser = argparse.ArgumentParser(description='Import JSON files to Master.xlsx')
    parser.add_argument('--base-dir', default='data', help='Base directory to scan (default: data)')
    parser.add_argument('--master', default='data/Master.xlsx', help='Master Excel file path (default: data/Master.xlsx)')
    
    args = parser.parse_args()
    
    result = process_json_files(args.base_dir, args.master)
    
    # Exit code: 0 if successful, 1 if errors
    sys.exit(0 if not result['errors'] else 1)
