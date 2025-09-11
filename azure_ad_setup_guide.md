# Azure AD App Registration Setup Guide

This guide walks you through setting up Azure AD app registration for modern SharePoint authentication.

## Prerequisites

- Azure AD tenant admin permissions (or delegated permissions to register apps)
- SharePoint Online access
- Python environment with required packages

## Step 1: Register Application in Azure AD

1. **Navigate to Azure Portal**
   - Go to [https://portal.azure.com](https://portal.azure.com)
   - Sign in with your organizational account

2. **Access Azure Active Directory**
   - Click on "Azure Active Directory" in the left navigation
   - Or search for "Azure Active Directory" in the search bar

3. **Register New Application**
   - Click on "App registrations" in the left menu
   - Click "New registration"
   - Fill in the application details:
     - **Name**: Give your app a descriptive name (e.g., "SharePoint Python Client")
     - **Supported account types**: Select "Accounts in this organizational directory only"
     - **Redirect URI**: Leave blank for now (or set to `http://localhost` if needed)
   - Click "Register"

4. **Note Important IDs**
   - After registration, you'll see the app overview page
   - **Copy and save these values**:
     - Application (client) ID
     - Directory (tenant) ID
     - Object ID

## Step 2: Create Client Secret

1. **Generate Client Secret**
   - In your app registration, go to "Certificates & secrets"
   - Click "New client secret"
   - Add a description (e.g., "Python Script Secret")
   - Choose expiration period (recommend 24 months max for security)
   - Click "Add"

2. **Copy Secret Value**
   - **IMPORTANT**: Copy the secret value immediately
   - You won't be able to see it again after leaving this page
   - Store it securely

## Step 3: Configure API Permissions

### For Microsoft Graph API (Recommended)

1. **Add Microsoft Graph Permissions**
   - Go to "API permissions"
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Application permissions" (for app-only access)

2. **Select Required Permissions**
   - **Sites.Read.All**: Read all site collections
   - **Sites.ReadWrite.All**: Read and write all site collections
   - **Files.Read.All**: Read all files
   - **Files.ReadWrite.All**: Read and write all files
   - Click "Add permissions"

3. **Grant Admin Consent**
   - Click "Grant admin consent for [Your Organization]"
   - Confirm by clicking "Yes"
   - Status should show "Granted for [Your Organization]"

### For SharePoint Online (Alternative)

1. **Add SharePoint Permissions**
   - Click "Add a permission"
   - Select "SharePoint"
   - Choose "Application permissions"

2. **Select Required Permissions**
   - **Sites.FullControl.All**: Full control of all site collections
   - **Sites.Read.All**: Read all site collections
   - Click "Add permissions"

3. **Grant Admin Consent**
   - Click "Grant admin consent for [Your Organization]"
   - Confirm by clicking "Yes"

## Step 4: Install Required Python Packages

```bash
# For MSAL + Microsoft Graph approach
pip install msal requests pandas openpyxl

# For Office365-REST-Python-Client approach
pip install Office365-REST-Python-Client pandas openpyxl

# For direct REST API calls
pip install requests pandas openpyxl
```

## Step 5: Test Authentication

Create a test script to verify your setup:

```python
import msal
import requests

# Your app registration details
CLIENT_ID = "your-client-id-here"
CLIENT_SECRET = "your-client-secret-here"
TENANT_ID = "your-tenant-id-here"

# Create MSAL app
app = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

# Get access token
result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

if "access_token" in result:
    print("✅ Authentication successful!")
    
    # Test API call
    headers = {"Authorization": f"Bearer {result['access_token']}"}
    response = requests.get("https://graph.microsoft.com/v1.0/sites", headers=headers)
    
    if response.status_code == 200:
        print("✅ API call successful!")
        sites = response.json()
        print(f"Found {len(sites.get('value', []))} sites")
    else:
        print(f"❌ API call failed: {response.status_code}")
else:
    print(f"❌ Authentication failed: {result.get('error_description', 'Unknown error')}")
```

## Step 6: Security Best Practices

1. **Store Credentials Securely**
   ```python
   import os
   
   # Use environment variables
   CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
   CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
   TENANT_ID = os.getenv('AZURE_TENANT_ID')
   ```

2. **Use Certificates Instead of Secrets** (Advanced)
   - Generate a certificate
   - Upload public key to Azure AD
   - Use private key in your application

3. **Principle of Least Privilege**
   - Only grant minimum required permissions
   - Review permissions regularly
   - Use site-specific permissions when possible

## Troubleshooting Common Issues

### Issue: "AADSTS70011: The provided value for the input parameter 'scope' is not valid"
**Solution**: Make sure you're using the correct scope format:
- For Microsoft Graph: `https://graph.microsoft.com/.default`
- For SharePoint: `https://yourtenant.sharepoint.com/.default`

### Issue: "AADSTS65001: The user or administrator has not consented to use the application"
**Solution**: Ensure admin consent has been granted for the required permissions.

### Issue: "Access denied" when calling SharePoint APIs
**Solution**: 
- Verify the app has the correct permissions
- Check that admin consent has been granted
- Ensure you're using the correct site URL format

### Issue: "The remote server returned an error: (401) Unauthorized"
**Solution**:
- Verify the access token is valid and not expired
- Check that the token has the required scopes
- Ensure the API endpoint URL is correct

## Alternative: Delegated Permissions (User Context)

If you need user context instead of app-only access:

1. **Use Delegated Permissions**
   - In API permissions, choose "Delegated permissions" instead of "Application permissions"
   - Select appropriate delegated permissions (e.g., `Sites.Read.All`, `Files.ReadWrite.All`)

2. **Implement User Authentication Flow**
   ```python
   # Use MSAL with user authentication
   result = app.acquire_token_interactive(scopes=["https://graph.microsoft.com/Sites.Read.All"])
   ```

## Next Steps

1. Choose one of the provided Python examples based on your needs
2. Replace the placeholder credentials with your actual values
3. Test with a small SharePoint site first
4. Implement error handling and logging for production use
5. Consider implementing token caching for better performance

## Useful Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
- [SharePoint REST API Reference](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Office365-REST-Python-Client Documentation](https://github.com/vgrem/Office365-REST-Python-Client)