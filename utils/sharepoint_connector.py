# utils/sharepoint_connector.py

import os
import requests
from msal import ConfidentialClientApplication
from typing import List, Dict, Optional, Tuple
import io
from pathlib import Path


class SharePointConnector:
    """
    Microsoft Graph API connector for SharePoint Online operations.
    """
    
    def __init__(self, client_id: str, client_secret: str, tenant_id: str, site_id: str):
        """
        Initialize SharePoint connector.
        
        Args:
            client_id: Azure AD Application (client) ID
            client_secret: Azure AD Application client secret
            tenant_id: Azure AD Tenant ID
            site_id: SharePoint site ID
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.site_id = site_id
        self.token = None
    
    def authenticate(self) -> str:
        """
        Authenticates with Microsoft Graph API and returns access token.
        
        Returns:
            str: Access token
        
        Raises:
            Exception: If authentication fails
        """
        app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "access_token" in result:
            self.token = result["access_token"]
            return self.token
        else:
            raise Exception(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
    
    def _get_headers(self) -> Dict[str, str]:
        """Returns authorization headers."""
        if not self.token:
            self.authenticate()
        return {"Authorization": f"Bearer {self.token}"}
    
    def list_folders(self, parent_path: str = "") -> List[Dict]:
        """
        Lists folders in SharePoint directory.
        
        Args:
            parent_path: Parent folder path (e.g., "/Documents/Data")
        
        Returns:
            List[Dict]: List of folder metadata
        """
        if parent_path:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:/{parent_path}:/children"
        else:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root/children"
        
        response = requests.get(url, headers=self._get_headers())
        
        if response.status_code == 200:
            items = response.json().get('value', [])
            # Filter folders only
            folders = [item for item in items if 'folder' in item]
            return folders
        else:
            raise Exception(f"Failed to list folders: {response.status_code} - {response.text}")
    
    def list_files(self, folder_path: str, extension: str = ".pdf") -> List[Dict]:
        """
        Lists files in a SharePoint folder.
        
        Args:
            folder_path: Folder path (e.g., "/Documents/HH1309")
            extension: File extension filter (default: .pdf)
        
        Returns:
            List[Dict]: List of file metadata
        """
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:/{folder_path}:/children"
        
        response = requests.get(url, headers=self._get_headers())
        
        if response.status_code == 200:
            items = response.json().get('value', [])
            # Filter files with specific extension
            files = [
                item for item in items 
                if 'file' in item and item['name'].lower().endswith(extension.lower())
            ]
            return files
        else:
            raise Exception(f"Failed to list files: {response.status_code} - {response.text}")
    
    def download_file(self, file_path: str) -> bytes:
        """
        Downloads a file from SharePoint.
        
        Args:
            file_path: Full path to file (e.g., "/Documents/HH1309/PO1.pdf")
        
        Returns:
            bytes: File content
        """
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:/{file_path}:/content"
        
        response = requests.get(url, headers=self._get_headers())
        
        if response.status_code == 200:
            return response.content
        else:
            raise Exception(f"Failed to download file: {response.status_code} - {response.text}")
    
    def upload_file(self, file_path: str, file_content: bytes, overwrite: bool = True) -> Dict:
        """
        Uploads a file to SharePoint.
        
        Args:
            file_path: Destination path in SharePoint (e.g., "/Documents/Master.xlsx")
            file_content: File content as bytes
            overwrite: Whether to overwrite existing file
        
        Returns:
            Dict: Upload result metadata
        """
        # For files < 4MB, use simple upload
        if len(file_content) < 4 * 1024 * 1024:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:/{file_path}:/content"
            
            headers = self._get_headers()
            headers['Content-Type'] = 'application/octet-stream'
            
            response = requests.put(url, headers=headers, data=file_content)
            
            if response.status_code in [200, 201]:
                return response.json()
            else:
                raise Exception(f"Failed to upload file: {response.status_code} - {response.text}")
        else:
            # For larger files, use upload session (not implemented in MVP)
            raise NotImplementedError("Large file upload (>4MB) requires upload session")
    
    def scan_saha_folders(self, base_folder: str = "/Documents") -> List[Tuple[str, List[Dict]]]:
        """
        Scans SharePoint for SahaID folders and their PDF files.
        
        Args:
            base_folder: Base folder path in SharePoint
        
        Returns:
            List[Tuple[str, List[Dict]]]: List of (saha_id, pdf_files) tuples
        """
        results = []
        
        # List all folders
        folders = self.list_folders(base_folder)
        
        # Scan all folders (no prefix restriction)
        saha_folders = folders
        
        for folder in saha_folders:
            saha_id = folder['name']
            folder_path = f"{base_folder}/{saha_id}"
            
            # List PDF files in this folder
            try:
                pdf_files = self.list_files(folder_path, extension=".pdf")
                if pdf_files:
                    results.append((saha_id, pdf_files))
            except Exception as e:
                print(f"Warning: Failed to list files in {folder_path}: {str(e)}")
        
        return results
    
    def download_saha_pdfs(self, base_folder: str = "/Documents", 
                          local_base_dir: str = "data") -> List[Tuple[str, str]]:
        """
        Downloads all PDFs from SahaID folders to local directory.
        
        Args:
            base_folder: SharePoint base folder
            local_base_dir: Local base directory
        
        Returns:
            List[Tuple[str, str]]: List of (saha_id, local_file_path) tuples
        """
        downloaded = []
        
        # Scan SharePoint folders
        saha_data = self.scan_saha_folders(base_folder)
        
        for saha_id, pdf_files in saha_data:
            # Create local directory
            local_saha_dir = Path(local_base_dir) / saha_id
            local_saha_dir.mkdir(parents=True, exist_ok=True)
            
            for pdf_file in pdf_files:
                try:
                    # Download file
                    file_name = pdf_file['name']
                    sharepoint_path = f"{base_folder}/{saha_id}/{file_name}"
                    
                    print(f"Downloading: {sharepoint_path}")
                    file_content = self.download_file(sharepoint_path)
                    
                    # Save locally
                    local_file_path = local_saha_dir / file_name
                    with open(local_file_path, 'wb') as f:
                        f.write(file_content)
                    
                    downloaded.append((saha_id, str(local_file_path)))
                    print(f"✅ Downloaded: {local_file_path}")
                
                except Exception as e:
                    print(f"❌ Failed to download {pdf_file['name']}: {str(e)}")
        
        return downloaded


def get_sharepoint_connector_from_env() -> SharePointConnector:
    """
    Creates SharePointConnector from environment variables.
    
    Environment variables required:
    - SHAREPOINT_CLIENT_ID
    - SHAREPOINT_CLIENT_SECRET
    - SHAREPOINT_TENANT_ID
    - SHAREPOINT_SITE_ID
    
    Returns:
        SharePointConnector: Initialized connector
    """
    client_id = os.getenv('SHAREPOINT_CLIENT_ID')
    client_secret = os.getenv('SHAREPOINT_CLIENT_SECRET')
    tenant_id = os.getenv('SHAREPOINT_TENANT_ID')
    site_id = os.getenv('SHAREPOINT_SITE_ID')
    
    if not all([client_id, client_secret, tenant_id, site_id]):
        raise ValueError("Missing SharePoint credentials in environment variables")
    
    return SharePointConnector(client_id, client_secret, tenant_id, site_id)
