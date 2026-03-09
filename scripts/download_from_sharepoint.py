#!/usr/bin/env python3
"""
Download PDFs from SharePoint to local data/ directory.

Usage:
    python scripts/download_from_sharepoint.py

Environment variables required:
    SHAREPOINT_CLIENT_ID
    SHAREPOINT_CLIENT_SECRET
    SHAREPOINT_TENANT_ID
    SHAREPOINT_SITE_ID
    SHAREPOINT_BASE_FOLDER (optional, default: /Documents)
"""

import os
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils.sharepoint_connector import get_sharepoint_connector_from_env


def main():
    print("🔄 Starting SharePoint download process...")
    
    try:
        # Get SharePoint connector
        sp = get_sharepoint_connector_from_env()
        
        # Get base folder from env or use default
        base_folder = os.getenv('SHAREPOINT_BASE_FOLDER', '/Documents')
        local_dir = 'data'
        
        print(f"📁 SharePoint base folder: {base_folder}")
        print(f"💾 Local directory: {local_dir}")
        
        # Download PDFs
        print("\n🌐 Authenticating with SharePoint...")
        sp.authenticate()
        print("✅ Authentication successful")
        
        print(f"\n📥 Scanning and downloading PDFs from {base_folder}...")
        downloaded = sp.download_saha_pdfs(base_folder, local_dir)
        
        # Summary
        print("\n" + "="*60)
        print(f"✅ Download completed!")
        print(f"📊 Total files downloaded: {len(downloaded)}")
        
        if downloaded:
            print("\nDownloaded files:")
            for saha_id, file_path in downloaded:
                print(f"  - {saha_id}: {Path(file_path).name}")
        else:
            print("\n⚠️  No new PDF files found in SharePoint")
        
        print("="*60)
        
    except ValueError as e:
        print(f"❌ Configuration error: {str(e)}")
        print("\nRequired environment variables:")
        print("  - SHAREPOINT_CLIENT_ID")
        print("  - SHAREPOINT_CLIENT_SECRET")
        print("  - SHAREPOINT_TENANT_ID")
        print("  - SHAREPOINT_SITE_ID")
        sys.exit(1)
    
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
