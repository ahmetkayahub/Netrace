#!/usr/bin/env python3
"""
Upload Master.xlsx to SharePoint.

Usage:
    python scripts/upload_to_sharepoint.py

Environment variables required:
    SHAREPOINT_CLIENT_ID
    SHAREPOINT_CLIENT_SECRET
    SHAREPOINT_TENANT_ID
    SHAREPOINT_SITE_ID
    SHAREPOINT_MASTER_PATH (optional, default: /Documents/Master.xlsx)
"""

import os
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils.sharepoint_connector import get_sharepoint_connector_from_env


def main():
    print("🔄 Starting Master.xlsx upload to SharePoint...")
    
    try:
        # Get SharePoint connector
        sp = get_sharepoint_connector_from_env()
        
        # Get paths
        local_master_path = 'data/Master.xlsx'
        sharepoint_path = os.getenv('SHAREPOINT_MASTER_PATH', '/Documents/Master.xlsx')
        
        # Check if local file exists
        if not Path(local_master_path).exists():
            print(f"❌ Local Master.xlsx not found: {local_master_path}")
            sys.exit(1)
        
        print(f"📁 Local file: {local_master_path}")
        print(f"☁️  SharePoint path: {sharepoint_path}")
        
        # Read local file
        print("\n📖 Reading local Master.xlsx...")
        with open(local_master_path, 'rb') as f:
            file_content = f.read()
        
        file_size_mb = len(file_content) / (1024 * 1024)
        print(f"📊 File size: {file_size_mb:.2f} MB")
        
        # Authenticate
        print("\n🌐 Authenticating with SharePoint...")
        sp.authenticate()
        print("✅ Authentication successful")
        
        # Upload
        print(f"\n📤 Uploading to SharePoint...")
        result = sp.upload_file(sharepoint_path, file_content, overwrite=True)
        
        print("\n" + "="*60)
        print("✅ Upload completed successfully!")
        print(f"📊 File: {result.get('name', 'Master.xlsx')}")
        print(f"🔗 Web URL: {result.get('webUrl', 'N/A')}")
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
