import requests
from datetime import datetime
import json
import msal

def search_sharepoint_files():
    # Azure AD app registration details
    client_id = input("Enter your Azure AD application client ID: ")
    tenant_id = input("Enter your Azure AD tenant ID: ")
    
    # Initialize MSAL app
    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    
    # Get token
    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_interactive(scopes)
    
    if "access_token" in result:
        # Search parameters
        site = "jjag.sharepoint.com"
        relative_path = "/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/2025/05 - May"
        search_pattern = "CS Flex Weekly Service Delivery Report"
        
        # Construct the search query
        headers = {
            'Authorization': f'Bearer {result["access_token"]}',
            'Content-Type': 'application/json'
        }
        
        search_query = {
            "requests": [{
                "entityTypes": ["driveItem"],
                "query": {
                    "queryString": f'site:"{site}" path:"{relative_path}" filename:"{search_pattern}*.xlsx"'
                }
            }]
        }
        
        # Make the search request
        response = requests.post(
            'https://graph.microsoft.com/v1.0/search/query',
            headers=headers,
            json=search_query
        )
        
        if response.status_code == 200:
            data = response.json()
            hits = data.get('value', [{}])[0].get('hitsContainers', [{}])[0].get('hits', [])
            
            if not hits:
                print("No matching files found.")
                return
            
            print(f"\nFound {len(hits)} matching files:\n")
            for hit in hits:
                resource = hit.get('resource', {})
                print(f"File: {resource.get('name')}")
                print(f"Path: {resource.get('webUrl')}")
                print("---")
        else:
            print(f"Error: {response.status_code}")
            print(response.text)
    else:
        print(f"Error getting token: {result.get('error_description', 'Unknown error')}")

if __name__ == "__main__":
    search_sharepoint_files() 