"""
SharePoint App-Only Authentication Example
This is an older method that may still work in some organizations.
Note: This method is being deprecated in favor of Azure AD app registration.
"""

from office365.runtime.auth.providers.sharepoint_provider import SharePointAuthenticationProvider
from office365.sharepoint.client_context import ClientContext
import requests
import base64
import json
from urllib.parse import urlencode
import xml.etree.ElementTree as ET

class SharePointAppOnlyAuth:
    def __init__(self, site_url: str, client_id: str, client_secret: str):
        """
        Initialize SharePoint App-Only authentication.
        
        Args:
            site_url: SharePoint site URL
            client_id: SharePoint App Client ID (from appregnew.aspx)
            client_secret: SharePoint App Client Secret (from appregnew.aspx)
        """
        self.site_url = site_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.access_token = None
        
        # Extract tenant and realm information from site URL
        self.tenant_name = site_url.split('//')[1].split('.')[0]
        self.realm = self._get_realm()
    
    def _get_realm(self) -> str:
        """
        Get the SharePoint realm (tenant ID) from the site.
        """
        try:
            # Try to get realm from SharePoint
            realm_url = f"{self.site_url}/_vti_bin/client.svc"
            response = requests.get(realm_url)
            
            if 'WWW-Authenticate' in response.headers:
                auth_header = response.headers['WWW-Authenticate']
                if 'realm=' in auth_header:
                    realm = auth_header.split('realm="')[1].split('"')[0]
                    return realm
            
            # Fallback: construct realm from tenant name
            return f"{self.tenant_name}.sharepoint.com"
            
        except Exception as e:
            print(f"Warning: Could not determine realm automatically: {e}")
            return f"{self.tenant_name}.sharepoint.com"
    
    def get_access_token(self) -> str:
        """
        Get access token using SharePoint App-Only authentication.
        """
        try:
            # SharePoint STS endpoint
            sts_url = f"https://accounts.accesscontrol.windows.net/{self.realm}/tokens/OAuth/2"
            
            # Prepare the request
            principal = f"{self.client_id}@{self.realm}"
            resource = f"00000003-0000-0ff1-ce00-000000000000/{self.site_url.split('//')[1].split('/')[0]}@{self.realm}"
            
            # Request body
            body = {
                'grant_type': 'client_credentials',
                'client_id': principal,
                'client_secret': self.client_secret,
                'resource': resource
            }
            
            # Make the request
            response = requests.post(
                sts_url,
                data=urlencode(body),
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                print("Successfully obtained access token using App-Only authentication")
                return self.access_token
            else:
                print(f"Failed to get access token: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            print(f"Error getting access token: {str(e)}")
            return None
    
    def create_client_context(self) -> ClientContext:
        """
        Create SharePoint client context with App-Only authentication.
        """
        if not self.access_token:
            self.get_access_token()
        
        if self.access_token:
            # Create authentication provider
            auth_provider = SharePointAuthenticationProvider(self.access_token)
            
            # Create client context
            ctx = ClientContext(self.site_url)
            ctx._auth_provider = auth_provider
            
            return ctx
        
        return None

# Alternative implementation using direct REST API calls
class SharePointRestClient:
    def __init__(self, site_url: str, access_token: str):
        """
        Initialize SharePoint REST client with access token.
        
        Args:
            site_url: SharePoint site URL
            access_token: OAuth access token
        """
        self.site_url = site_url
        self.access_token = access_token
        self.api_base = f"{site_url}/_api"
    
    def get_headers(self) -> dict:
        """Get headers for REST API requests."""
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }
    
    def get_web_info(self) -> dict:
        """Get SharePoint web information."""
        url = f"{self.api_base}/web"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting web info: {response.status_code} - {response.text}")
            return {}
    
    def get_lists(self) -> dict:
        """Get all lists in the SharePoint site."""
        url = f"{self.api_base}/web/lists"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting lists: {response.status_code} - {response.text}")
            return {}
    
    def get_list_items(self, list_title: str) -> dict:
        """Get items from a SharePoint list."""
        url = f"{self.api_base}/web/lists/getbytitle('{list_title}')/items"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting list items: {response.status_code} - {response.text}")
            return {}
    
    def get_files_in_library(self, library_name: str) -> dict:
        """Get files in a document library."""
        url = f"{self.api_base}/web/lists/getbytitle('{library_name}')/items?$expand=File"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error getting files: {response.status_code} - {response.text}")
            return {}
    
    def download_file(self, file_server_relative_url: str) -> bytes:
        """Download a file from SharePoint."""
        url = f"{self.site_url}/_api/web/getfilebyserverrelativeurl('{file_server_relative_url}')/$value"
        response = requests.get(url, headers=self.get_headers())
        
        if response.status_code == 200:
            return response.content
        else:
            print(f"Error downloading file: {response.status_code} - {response.text}")
            return None

# Setup instructions for SharePoint App-Only authentication
def print_app_only_setup_instructions():
    """
    Print instructions for setting up SharePoint App-Only authentication.
    """
    instructions = """
    SharePoint App-Only Authentication Setup Instructions:
    
    WARNING: This method is being deprecated. Use Azure AD app registration instead.
    
    1. Register a new SharePoint App:
       - Navigate to: https://yourtenant.sharepoint.com/_layouts/15/appregnew.aspx
       - Click "Generate" for both Client Id and Client Secret
       - Set Title: Your app name
       - App Domain: localhost (or your domain)
       - Redirect URI: https://localhost (or your redirect URI)
       - Click "Create"
       - Save the Client Id and Client Secret
    
    2. Grant permissions to the app:
       - Navigate to: https://yourtenant.sharepoint.com/_layouts/15/appinv.aspx
       - Enter the Client Id from step 1 and click "Lookup"
       - In the Permission Request XML field, enter:
         <AppPermissionRequests AllowAppOnlyPolicy="true">
           <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
         </AppPermissionRequests>
       - Click "Create"
       - Click "Trust It" when prompted
    
    3. Use the Client Id and Client Secret in your Python script
    
    Note: You may need tenant admin permissions to complete these steps.
    """
    print(instructions)

# Example usage
def main():
    # Print setup instructions
    print_app_only_setup_instructions()
    
    # Configuration - replace with your actual values
    SITE_URL = "https://yourtenant.sharepoint.com/sites/yoursite"
    CLIENT_ID = "your-app-client-id-here"  # From appregnew.aspx
    CLIENT_SECRET = "your-app-client-secret-here"  # From appregnew.aspx
    
    # Try App-Only authentication
    app_auth = SharePointAppOnlyAuth(SITE_URL, CLIENT_ID, CLIENT_SECRET)
    access_token = app_auth.get_access_token()
    
    if access_token:
        print("App-Only authentication successful!")
        
        # Use REST client
        rest_client = SharePointRestClient(SITE_URL, access_token)
        
        # Get web information
        web_info = rest_client.get_web_info()
        if web_info and 'd' in web_info:
            print(f"Site Title: {web_info['d']['Title']}")
            print(f"Site URL: {web_info['d']['Url']}")
        
        # Get lists
        lists_info = rest_client.get_lists()
        if lists_info and 'd' in lists_info:
            print("\nAvailable Lists:")
            for lst in lists_info['d']['results']:
                print(f"- {lst['Title']} (Type: {lst['BaseTemplate']})")
        
        # Get files from Documents library (if it exists)
        try:
            files_info = rest_client.get_files_in_library("Documents")
            if files_info and 'd' in files_info:
                print("\nFiles in Documents library:")
                for item in files_info['d']['results']:
                    if 'File' in item and item['File']:
                        print(f"- {item['File']['Name']}")
        except Exception as e:
            print(f"Could not access Documents library: {e}")
    
    else:
        print("App-Only authentication failed. This method may not be available in your organization.")
        print("Try using Azure AD app registration with MSAL instead.")

if __name__ == "__main__":
    main()