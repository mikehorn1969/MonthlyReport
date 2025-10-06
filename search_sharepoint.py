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
        client_secret = CLIENT_SECRET

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
        site_response = requests.get(site_url, headers=headers)
        
        if site_response.status_code == 200:            
            site_id = site_response.json()['id']            
        else:
            print(f"Error getting site info: {site_response.status_code}")
            print(site_response.text)
            return []
        
        # Now try to get the list by list display name
        list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{urllib.parse.quote(list_name)}/items"
                
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


def mark_file_as_processed(access_token, item_id, site_name="jjag.sharepoint.com", list_name="Service Provider Uploads"):
    """Update a SharePoint list item to mark it as processed"""
    try:
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Get site ID first
        site_info_url = f"https://graph.microsoft.com/v1.0/sites/{site_name}:/sites/InternalTeam:"
        site_response = requests.get(site_info_url, headers=headers)
        
        if site_response.status_code != 200:
            print(f"Failed to get site info: {site_response.status_code}")
            return False
            
        site_id = site_response.json()['id']
        
        # Update the list item
        update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/items/{item_id}/fields"
        
        update_data = {
            "Monthlyreportprocessed": True
        }
        
        response = requests.patch(update_url, headers=headers, json=update_data)
        
        if response.status_code in [200, 201, 204]:
            print(f"Successfully marked item {item_id} as processed")
            return True
        else:
            print(f"Failed to mark item as processed: {response.status_code} - {response.text}")
            return False
            
    except Exception as e:
        print(f"Error marking file as processed: {str(e)}")
        return False