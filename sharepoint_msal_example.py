"""
SharePoint Excel File Access using MSAL and Microsoft Graph API
This example shows how to authenticate with modern OAuth2 and access Excel files from SharePoint.
"""

import msal
import requests
import json
from typing import Optional, Dict, Any

class SharePointGraphClient:
    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        """
        Initialize the SharePoint Graph client with Azure AD app credentials.
        
        Args:
            client_id: Azure AD Application (client) ID
            client_secret: Azure AD Application client secret
            tenant_id: Azure AD Directory (tenant) ID
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        
        # Microsoft Graph API endpoint
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"
        
        # Create MSAL confidential client application
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
    
    def get_access_token(self) -> Optional[str]:
        """
        Acquire access token using client credentials flow (app-only authentication).
        """
        # Define the scopes for Microsoft Graph
        scopes = ["https://graph.microsoft.com/.default"]
        
        try:
            # Acquire token using client credentials
            result = self.app.acquire_token_silent(scopes, account=None)
            
            if not result:
                result = self.app.acquire_token_for_client(scopes=scopes)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                return self.access_token
            else:
                print(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return None
                
        except Exception as e:
            print(f"Error acquiring access token: {str(e)}")
            return None
    
    def get_headers(self) -> Dict[str, str]:
        """Get headers for API requests."""
        if not self.access_token:
            self.get_access_token()
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
    
    def get_sharepoint_sites(self, site_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Get SharePoint sites accessible to the application.
        
        Args:
            site_name: Optional filter by site name
        """
        url = f"{self.graph_endpoint}/sites"
        if site_name:
            url += f"?search={site_name}"
        
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting sites: {response.status_code} - {response.text}")
            return {}
    
    def get_site_by_url(self, site_url: str) -> Dict[str, Any]:
        """
        Get a specific SharePoint site by its URL.
        
        Args:
            site_url: Full SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)
        """
        # Extract hostname and site path from URL
        from urllib.parse import urlparse
        parsed = urlparse(site_url)
        hostname = parsed.netloc
        site_path = parsed.path
        
        url = f"{self.graph_endpoint}/sites/{hostname}:{site_path}"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting site: {response.status_code} - {response.text}")
            return {}
    
    def get_site_drives(self, site_id: str) -> Dict[str, Any]:
        """
        Get document libraries (drives) in a SharePoint site.
        
        Args:
            site_id: SharePoint site ID
        """
        url = f"{self.graph_endpoint}/sites/{site_id}/drives"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting drives: {response.status_code} - {response.text}")
            return {}
    
    def get_drive_items(self, site_id: str, drive_id: str, folder_path: str = "") -> Dict[str, Any]:
        """
        Get items in a document library or folder.
        
        Args:
            site_id: SharePoint site ID
            drive_id: Document library (drive) ID
            folder_path: Optional folder path within the drive
        """
        if folder_path:
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root:/{folder_path}:/children"
        else:
            url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/root/children"
        
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting drive items: {response.status_code} - {response.text}")
            return {}
    
    def download_excel_file(self, site_id: str, drive_id: str, item_id: str, local_path: str) -> bool:
        """
        Download an Excel file from SharePoint.
        
        Args:
            site_id: SharePoint site ID
            drive_id: Document library (drive) ID
            item_id: File item ID
            local_path: Local path to save the file
        """
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/items/{item_id}/content"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            with open(local_path, 'wb') as f:
                f.write(response.content)
            print(f"File downloaded successfully to {local_path}")
            return True
        else:
            print(f"Error downloading file: {response.status_code} - {response.text}")
            return False
    
    def get_excel_worksheets(self, site_id: str, drive_id: str, item_id: str) -> Dict[str, Any]:
        """
        Get worksheets from an Excel file in SharePoint.
        
        Args:
            site_id: SharePoint site ID
            drive_id: Document library (drive) ID
            item_id: Excel file item ID
        """
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/items/{item_id}/workbook/worksheets"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting worksheets: {response.status_code} - {response.text}")
            return {}
    
    def get_excel_data(self, site_id: str, drive_id: str, item_id: str, worksheet_name: str, range_address: str = "A1:Z1000") -> Dict[str, Any]:
        """
        Get data from an Excel worksheet in SharePoint.
        
        Args:
            site_id: SharePoint site ID
            drive_id: Document library (drive) ID
            item_id: Excel file item ID
            worksheet_name: Name of the worksheet
            range_address: Cell range to retrieve (e.g., "A1:Z1000")
        """
        url = f"{self.graph_endpoint}/sites/{site_id}/drives/{drive_id}/items/{item_id}/workbook/worksheets('{worksheet_name}')/range(address='{range_address}')"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting Excel data: {response.status_code} - {response.text}")
            return {}

# Example usage
def main():
    # Azure AD app credentials (replace with your actual values)
    CLIENT_ID = "your-client-id-here"
    CLIENT_SECRET = "your-client-secret-here"
    TENANT_ID = "your-tenant-id-here"
    
    # SharePoint site URL
    SITE_URL = "https://yourtenant.sharepoint.com/sites/yoursite"
    
    # Initialize the client
    client = SharePointGraphClient(CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    
    # Get access token
    if not client.get_access_token():
        print("Failed to authenticate")
        return
    
    print("Authentication successful!")
    
    # Get the SharePoint site
    site_info = client.get_site_by_url(SITE_URL)
    if not site_info:
        print("Failed to get site information")
        return
    
    site_id = site_info["id"]
    print(f"Site ID: {site_id}")
    
    # Get document libraries
    drives = client.get_site_drives(site_id)
    if not drives or "value" not in drives:
        print("No document libraries found")
        return
    
    # Find the first document library
    drive_id = drives["value"][0]["id"]
    print(f"Using drive ID: {drive_id}")
    
    # Get files in the document library
    items = client.get_drive_items(site_id, drive_id)
    if not items or "value" not in items:
        print("No files found")
        return
    
    # Find Excel files
    excel_files = [item for item in items["value"] if item["name"].endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("No Excel files found")
        return
    
    # Process the first Excel file
    excel_file = excel_files[0]
    file_name = excel_file["name"]
    item_id = excel_file["id"]
    
    print(f"Processing Excel file: {file_name}")
    
    # Option 1: Download the file
    client.download_excel_file(site_id, drive_id, item_id, f"./{file_name}")
    
    # Option 2: Read Excel data directly via Graph API
    worksheets = client.get_excel_worksheets(site_id, drive_id, item_id)
    if worksheets and "value" in worksheets:
        worksheet_name = worksheets["value"][0]["name"]
        excel_data = client.get_excel_data(site_id, drive_id, item_id, worksheet_name)
        
        if excel_data and "values" in excel_data:
            print(f"Excel data from worksheet '{worksheet_name}':")
            for row in excel_data["values"][:5]:  # Print first 5 rows
                print(row)

if __name__ == "__main__":
    main()