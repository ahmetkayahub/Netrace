#!/usr/bin/env python3
"""
Automated import script: Downloads PDFs from SharePoint, processes them, and uploads results.

This script is designed to run in GitHub Actions or as a scheduled task.

Usage:
    python scripts/auto_import.py
"""

import os
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils.sharepoint_connector import get_sharepoint_connector_from_env
from utils.importer import import_new_pdfs_to_master


def main():
    print("="*70)
    print("🚀 NETRACE AUTO IMPORT - Automated PDF Processing")
    print("="*70)
    
    # Configuration
    use_sharepoint = os.getenv('USE_SHAREPOINT', 'false').lower() == 'true'
    local_data_dir = 'data'
    master_path = 'data/Master.xlsx'
    state_path = 'data/import_state.json'
    
    try:
        # Step 1: Download from SharePoint (if enabled)
        if use_sharepoint:
            print("\n📥 STEP 1: Downloading PDFs from SharePoint...")
            print("-" * 70)
            
            try:
                sp = get_sharepoint_connector_from_env()
                base_folder = os.getenv('SHAREPOINT_BASE_FOLDER', '/Documents')
                
                sp.authenticate()
                print("✅ SharePoint authentication successful")
                
                downloaded = sp.download_saha_pdfs(base_folder, local_data_dir)
                print(f"✅ Downloaded {len(downloaded)} PDF file(s)")
                
            except Exception as e:
                print(f"⚠️  SharePoint download failed: {str(e)}")
                print("ℹ️  Continuing with local files...")
        else:
            print("\n📁 STEP 1: Using local files (SharePoint disabled)")
            print("-" * 70)
            print(f"ℹ️  Set USE_SHAREPOINT=true to enable SharePoint integration")
        
        # Step 2: Process PDFs
        print("\n⚙️  STEP 2: Processing PDFs and updating Master.xlsx...")
        print("-" * 70)
        
        result = import_new_pdfs_to_master(
            base_dir=local_data_dir,
            master_path=master_path,
            state_path=state_path
        )
        
        # Display results
        print(f"\n📊 Processing Results:")
        print(f"  • PDFs found: {result['total_pdfs_found']}")
        print(f"  • New PDFs processed: {result['new_pdfs_processed']}")
        print(f"  • Lines added to Excel: {result['total_lines_added']}")
        print(f"  • PDFs skipped (already processed): {result['skipped_pdfs']}")
        
        if result['errors']:
            print(f"\n⚠️  Errors encountered: {len(result['errors'])}")
            for error in result['errors'][:5]:  # Show first 5 errors
                print(f"  • {Path(error['file']).name}: {error['error']}")
        else:
            print("\n✅ All files processed successfully!")
        
        # Step 3: Upload to SharePoint (if enabled)
        if use_sharepoint and result['new_pdfs_processed'] > 0:
            print("\n📤 STEP 3: Uploading Master.xlsx to SharePoint...")
            print("-" * 70)
            
            try:
                sharepoint_path = os.getenv('SHAREPOINT_MASTER_PATH', '/Documents/Master.xlsx')
                
                # Read local Master.xlsx
                with open(master_path, 'rb') as f:
                    file_content = f.read()
                
                # Upload
                sp.upload_file(sharepoint_path, file_content, overwrite=True)
                print(f"✅ Master.xlsx uploaded to SharePoint: {sharepoint_path}")
                
            except Exception as e:
                print(f"❌ SharePoint upload failed: {str(e)}")
        else:
            print("\n📤 STEP 3: SharePoint upload skipped")
            print("-" * 70)
            if not use_sharepoint:
                print("ℹ️  SharePoint disabled")
            elif result['new_pdfs_processed'] == 0:
                print("ℹ️  No new data to upload")
        
        # Summary
        print("\n" + "="*70)
        print("✅ AUTO IMPORT COMPLETED")
        print("="*70)
        print(f"📊 Summary:")
        print(f"  • New PDFs processed: {result['new_pdfs_processed']}")
        print(f"  • Total lines added: {result['total_lines_added']}")
        print(f"  • Errors: {len(result['errors'])}")
        print("="*70)
        
        # Exit code based on errors
        if result['errors']:
            sys.exit(1)  # Exit with error if any files failed
        else:
            sys.exit(0)  # Success
        
    except Exception as e:
        print(f"\n❌ FATAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
