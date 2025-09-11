"""
SharePoint Excel File Access using Office365-REST-Python-Client
This example shows how to use the Office365-REST-Python-Client library with modern authentication.
"""

from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import os
import pandas as pd
from typing import List, Optional

class SharePointOffice365Client:
    def __init__(self, site_url: str, client_id: str, client_secret: str):
        """
        Initialize SharePoint client with modern authentication.
        
        Args:
            site_url: SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)
            client_id: Azure AD Application (client) ID
            client_secret: Azure AD Application client secret
        """
        self.site_url = site_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.ctx = None
        
    def authenticate(self) -> bool:
        """
        Authenticate using client credentials (app-only authentication).
        """
        try:
            # Create client credentials
            credentials = ClientCredential(self.client_id, self.client_secret)
            
            # Create SharePoint context with credentials
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
            
            # Test the connection by getting web properties
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            
            print(f"Successfully authenticated to SharePoint site: {web.title}")
            return True
            
        except Exception as e:
            print(f"Authentication failed: {str(e)}")
            return False
    
    def get_document_libraries(self) -> List[dict]:
        """
        Get all document libraries in the SharePoint site.
        """
        if not self.ctx:
            print("Not authenticated. Please call authenticate() first.")
            return []
        
        try:
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            
            doc_libraries = []
            for lst in lists:
                if lst.base_template == 101:  # Document Library template
                    doc_libraries.append({
                        'title': lst.title,
                        'id': lst.id,
                        'item_count': lst.item_count,
                        'server_relative_url': lst.default_view_url
                    })
            
            return doc_libraries
            
        except Exception as e:
            print(f"Error getting document libraries: {str(e)}")
            return []
    
    def get_files_in_library(self, library_name: str, folder_path: str = "") -> List[dict]:
        """
        Get files in a document library or specific folder.
        
        Args:
            library_name: Name of the document library
            folder_path: Optional folder path within the library
        """
        if not self.ctx:
            print("Not authenticated. Please call authenticate() first.")
            return []
        
        try:
            if folder_path:
                target_folder = self.ctx.web.get_folder_by_server_relative_url(
                    f"/sites/{self.site_url.split('/')[-1]}/{library_name}/{folder_path}"
                )
            else:
                target_folder = self.ctx.web.lists.get_by_title(library_name).root_folder
            
            files = target_folder.files
            self.ctx.load(files)
            self.ctx.execute_query()
            
            file_list = []
            for file in files:
                file_list.append({
                    'name': file.name,
                    'server_relative_url': file.server_relative_url,
                    'size': file.length,
                    'time_created': file.time_created.isoformat() if file.time_created else None,
                    'time_last_modified': file.time_last_modified.isoformat() if file.time_last_modified else None
                })
            
            return file_list
            
        except Exception as e:
            print(f"Error getting files: {str(e)}")
            return []
    
    def download_file(self, file_server_relative_url: str, local_path: str) -> bool:
        """
        Download a file from SharePoint.
        
        Args:
            file_server_relative_url: Server relative URL of the file
            local_path: Local path to save the file
        """
        if not self.ctx:
            print("Not authenticated. Please call authenticate() first.")
            return False
        
        try:
            # Create local directory if it doesn't exist
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            
            # Download file
            with open(local_path, 'wb') as local_file:
                file = self.ctx.web.get_file_by_server_relative_url(file_server_relative_url)
                file.download(local_file)
                self.ctx.execute_query()
            
            print(f"File downloaded successfully to: {local_path}")
            return True
            
        except Exception as e:
            print(f"Error downloading file: {str(e)}")
            return False
    
    def get_excel_files(self, library_name: str, folder_path: str = "") -> List[dict]:
        """
        Get all Excel files in a document library or folder.
        
        Args:
            library_name: Name of the document library
            folder_path: Optional folder path within the library
        """
        all_files = self.get_files_in_library(library_name, folder_path)
        excel_files = [f for f in all_files if f['name'].lower().endswith(('.xlsx', '.xls', '.xlsm'))]
        return excel_files
    
    def download_and_read_excel(self, file_server_relative_url: str, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
        """
        Download an Excel file from SharePoint and read it into a pandas DataFrame.
        
        Args:
            file_server_relative_url: Server relative URL of the Excel file
            sheet_name: Optional specific sheet name to read
        """
        if not self.ctx:
            print("Not authenticated. Please call authenticate() first.")
            return None
        
        try:
            # Create temporary file path
            file_name = os.path.basename(file_server_relative_url)
            temp_path = f"/tmp/{file_name}"
            
            # Download the file
            if not self.download_file(file_server_relative_url, temp_path):
                return None
            
            # Read Excel file into DataFrame
            if sheet_name:
                df = pd.read_excel(temp_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(temp_path)
            
            # Clean up temporary file
            os.remove(temp_path)
            
            print(f"Excel file read successfully. Shape: {df.shape}")
            return df
            
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return None
    
    def get_list_items(self, list_name: str, fields: Optional[List[str]] = None) -> List[dict]:
        """
        Get items from a SharePoint list.
        
        Args:
            list_name: Name of the SharePoint list
            fields: Optional list of field names to retrieve
        """
        if not self.ctx:
            print("Not authenticated. Please call authenticate() first.")
            return []
        
        try:
            sp_list = self.ctx.web.lists.get_by_title(list_name)
            items = sp_list.items
            
            if fields:
                self.ctx.load(items, fields)
            else:
                self.ctx.load(items)
            
            self.ctx.execute_query()
            
            item_list = []
            for item in items:
                item_dict = {}
                if fields:
                    for field in fields:
                        item_dict[field] = getattr(item, field, None)
                else:
                    # Get all properties
                    item_dict = item.properties
                
                item_list.append(item_dict)
            
            return item_list
            
        except Exception as e:
            print(f"Error getting list items: {str(e)}")
            return []

# Alternative authentication method using username/password (if allowed by organization)
class SharePointUserAuth:
    def __init__(self, site_url: str, username: str, password: str):
        """
        Initialize SharePoint client with user credentials.
        Note: This may not work if your organization has disabled legacy authentication.
        
        Args:
            site_url: SharePoint site URL
            username: SharePoint username
            password: SharePoint password
        """
        self.site_url = site_url
        self.username = username
        self.password = password
        self.ctx = None
    
    def authenticate(self) -> bool:
        """
        Authenticate using user credentials.
        """
        try:
            auth_context = AuthenticationContext(self.site_url)
            auth_context.acquire_token_for_user(self.username, self.password)
            
            self.ctx = ClientContext(self.site_url, auth_context)
            
            # Test connection
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            
            print(f"Successfully authenticated as user: {self.username}")
            return True
            
        except Exception as e:
            print(f"User authentication failed: {str(e)}")
            return False

# Example usage
def main():
    # Configuration - replace with your actual values
    SITE_URL = "https://yourtenant.sharepoint.com/sites/yoursite"
    CLIENT_ID = "your-client-id-here"
    CLIENT_SECRET = "your-client-secret-here"
    LIBRARY_NAME = "Documents"  # or your specific document library name
    
    # Initialize client
    client = SharePointOffice365Client(SITE_URL, CLIENT_ID, CLIENT_SECRET)
    
    # Authenticate
    if not client.authenticate():
        print("Authentication failed. Please check your credentials.")
        return
    
    # Get document libraries
    print("\nDocument Libraries:")
    libraries = client.get_document_libraries()
    for lib in libraries:
        print(f"- {lib['title']} (Items: {lib['item_count']})")
    
    # Get Excel files in the specified library
    print(f"\nExcel files in '{LIBRARY_NAME}' library:")
    excel_files = client.get_excel_files(LIBRARY_NAME)
    
    for excel_file in excel_files:
        print(f"- {excel_file['name']} (Size: {excel_file['size']} bytes)")
        
        # Download and read the Excel file
        df = client.download_and_read_excel(excel_file['server_relative_url'])
        if df is not None:
            print(f"  Columns: {list(df.columns)}")
            print(f"  First 3 rows:")
            print(df.head(3))
            print()

if __name__ == "__main__":
    main()