import requests
from datetime import datetime
import json
import msal
import urllib.parse

def get_sharepoint_access_token(client_id=None, tenant_id=None):
    """Get SharePoint access token using MSAL (exact same pattern as parse_reports.py)"""
    try:
        # Use config values or prompt if not provided
        if not client_id:
            try:
                from sharepoint_config import CLIENT_ID
                client_id = CLIENT_ID
            except:
                client_id = input("Enter your Azure AD application client ID: ")
        
        if not tenant_id:
            try:
                from sharepoint_config import TENANT_ID
                tenant_id = TENANT_ID
            except:
                tenant_id = input("Enter your Azure AD tenant ID: ")
        
        # Use the same scopes that work in parse_reports.py
        scopes = ["https://graph.microsoft.com/.default"]
        
        # Tech debt: Temporary workaround should be stored in Key Vault (same as parse_reports.py)
        client_secret = "fPl8Q~oEqBx.Mi0sfTJq2PQ-teDBeaHG6M5K4cKN"

        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=f"https://login.microsoftonline.com/{tenant_id}"
        )
       
        result = app.acquire_token_for_client(scopes=scopes)
        
        if "access_token" in result:
            print("Authentication successful!")
            return result["access_token"]
        else:
            print(f"Error getting token: {result.get('error_description', 'Unknown error')}")
            return None
            
    except Exception as e:
        print(f"Authentication error: {str(e)}")
        return None

def get_sharepoint_list_items(access_token, site_name, list_name, filter_query=None):
    """Get items from a SharePoint list using the same approach as parse_reports.py"""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }
    
    try:
        # Method 1: Try direct site access using hostname:/sites/sitename format
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_name}:/sites/InternalTeam"
        print(f"Trying site URL: {site_url}")
        site_response = requests.get(site_url, headers=headers)
        
        if site_response.status_code != 200:
            print(f"Method 1 failed: {site_response.status_code}")
            # Method 2: Try with root site access
            site_url = f"https://graph.microsoft.com/v1.0/sites/root"
            print(f"Trying root site URL: {site_url}")
            site_response = requests.get(site_url, headers=headers)
            
            if site_response.status_code != 200:
                print(f"Method 2 failed: {site_response.status_code}")
                # Method 3: Try searching for the site
                search_url = f"https://graph.microsoft.com/v1.0/sites?search=InternalTeam"
                print(f"Trying site search: {search_url}")
                search_response = requests.get(search_url, headers=headers)
                
                if search_response.status_code == 200:
                    sites = search_response.json().get('value', [])
                    if sites:
                        site_id = sites[0]['id']
                        print(f"Found site via search: {site_id}")
                    else:
                        print("No sites found via search")
                        return []
                else:
                    print(f"Site search failed: {search_response.status_code}")
                    print(search_response.text)
                    return []
            else:
                site_id = site_response.json()['id']
                print(f"Got root site ID: {site_id}")
        else:
            site_id = site_response.json()['id']
            print(f"Got site ID: {site_id}")
        
        # Now try to get the list
        # Method 1: Try by list display name
        list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{urllib.parse.quote(list_name)}/items"
        print(f"Trying list URL: {list_url}")
        
        # Add expand to get field values and filter if provided
        params = {
            'expand': 'fields'
        }
        
        if filter_query:
            params['filter'] = filter_query
        
        list_response = requests.get(list_url, headers=headers, params=params)
        
        if list_response.status_code == 200:
            return list_response.json().get('value', [])
        else:
            print(f"List access failed: {list_response.status_code}")
            print(list_response.text)
            
            # Method 2: Try to get all lists to see what's available
            lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
            lists_response = requests.get(lists_url, headers=headers)
            
            if lists_response.status_code == 200:
                lists = lists_response.json().get('value', [])
                print(f"Available lists in the site:")
                for lst in lists:
                    print(f"  - {lst.get('displayName')} (ID: {lst.get('id')})")
                
                # Try to find matching list
                for lst in lists:
                    if list_name.lower() in lst.get('displayName', '').lower():
                        print(f"Found matching list: {lst.get('displayName')}")
                        list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lst.get('id')}/items"
                        list_response = requests.get(list_url, headers=headers, params=params)
                        if list_response.status_code == 200:
                            return list_response.json().get('value', [])
            else:
                print(f"Could not get lists: {lists_response.status_code}")
            
            return []
            
    except Exception as e:
        print(f"Error getting SharePoint list items: {str(e)}")
        return []



def get_list_columns(access_token, site_name, list_name):
    """Get all column names from a SharePoint list (for debugging)"""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }
    
    # Get site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_name}"
    site_response = requests.get(site_url, headers=headers)
    
    if site_response.status_code != 200:
        print(f"Error getting site info: {site_response.status_code}")
        return []
    
    site_id = site_response.json()['id']
    
    # Get list columns
    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/columns"
    columns_response = requests.get(columns_url, headers=headers)
    
    if columns_response.status_code == 200:
        columns = columns_response.json().get('value', [])
        print(f"\nAvailable columns in '{list_name}' list:")
        print("=" * 50)
        for col in columns:
            print(f"- {col.get('name')} ({col.get('displayName')})")
        return columns
    else:
        print(f"Error getting columns: {columns_response.status_code}")
        return []

if __name__ == "__main__":
    # Main execution
    print("SharePoint List File Finder")
    print("=" * 40)
    print("This script queries the 'Service Provider Uploads' SharePoint list")
    print("to find report files using the 'path' and 'report filename' columns.\n")
    
    # Get authentication token once
    print("Authenticating with SharePoint...")
    access_token = get_sharepoint_access_token()
    
    if not access_token:
        print("❌ Authentication failed. Exiting.")
        exit(1)
    
    # Option to show available columns first
    # show_columns = input("Show available columns first? (y/n): ").lower().strip()
    
    # if show_columns == 'y':
    #     get_list_columns(access_token, "jjag.sharepoint.com", "Service Provider Uploads")
    #     print("\n" + "=" * 50)
    
    # Search for report files
    print("Searching for report files...")
    
    # Get list items
    site_name = "jjag.sharepoint.com"
    list_name = "Service Provider Uploads"
    
    print(f"Querying SharePoint list: {list_name}")
    
    # Get all items from the list
    list_items = get_sharepoint_list_items(access_token, site_name, list_name)
    
    if not list_items:
        print("No items found in the SharePoint list.")
    else:
        print(f"\nFound {len(list_items)} items in the list:")
        print("=" * 80)
        
        # Process and display items
        report_files = []
        
        for item in list_items:
            fields = item.get('fields', {})
            
            # Get the path and report filename columns (using correct SharePoint field names)
            path = fields.get('Path', 'N/A')
            filename = fields.get('Reportfilename', 'N/A')
            
            # Additional fields that might be useful
            title = fields.get('Title', 'N/A')
            created = fields.get('Created', 'N/A')
            modified = fields.get('Modified', 'N/A')
            
            print(f"Title: {title}")
            print(f"Path: {path}")
            print(f"Report Filename: {filename}")
            print(f"Created: {created}")
            print(f"Modified: {modified}")
            
            # Filter for files that look like reports
            if ('CS Flex Weekly Service Delivery Report' in str(filename) or 
                'Weekly Service Delivery' in str(filename) or
                '.xlsx' in str(filename)):
                
                report_files.append({
                    'title': title,
                    'path': path,
                    'filename': filename,
                    'created': created,
                    'modified': modified,
                    'full_url': f"{path}/{filename}" if path != 'N/A' and filename != 'N/A' else 'N/A'
                })
            
            print("-" * 40)
        
        if report_files:
            print(f"\n🎯 Found {len(report_files)} report files:")
            print("=" * 80)
            
            for i, report in enumerate(report_files, 1):
                print(f"{i}. {report['filename']}")
                print(f"   Path: {report['path']}")
                print(f"   Full URL: {report['full_url']}")
                print(f"   Modified: {report['modified']}")
                print()
            
            print(f"\n✅ Successfully found {len(report_files)} report files from SharePoint list!")
        else:
            print("\n❌ No report files found matching the criteria.")
            print("Available field names in the list items:")
            if list_items:
                sample_fields = list_items[0].get('fields', {}).keys()
                for field_name in sorted(sample_fields):
                    print(f"  - {field_name}")
                print("\nYou may need to adjust the field name mappings in the script.") 