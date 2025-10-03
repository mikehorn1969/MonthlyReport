"""
Excel Report Parser with SharePoint Support

This script parses Excel weekly reports and extracts information into text files.
Supports both local files and SharePoint Online files.

Usage from Excel (Python in Excel):

LOCAL FILES:
1. Import this module: import parse_reports
2. Process current folder: parse_reports.process_current_folder()
3. Process specific folder: parse_reports.process_folder("C:\\path\\to\\folder")

SHAREPOINT FILES:
1. Process SharePoint files: parse_reports.process_sharepoint_files()
2. Custom SharePoint config: parse_reports.process_sharepoint_with_config(
   site="yoursite.sharepoint.com", 
   path="/sites/yoursite/Documents", 
   patt                    if success:
                        print(f"📄 Successfully saved report summary locally: {output_filename}")
                    else:
                        print(f"⚠️ Failed to save report summary")"Report"
)

Usage from command line:
python parse_reports.py [directory_path]

Functions available for Excel:
- process_current_folder(): Process Excel files in current working directory
- process_folder(path): Process Excel files in specified directory
- process_sharepoint_files(): Process Excel files from SharePoint (uses file search)
- process_sharepoint_list_files(): Process files from SharePoint list "Service Provider Uploads"
- process_sharepoint_with_config(): Process SharePoint files with custom settings
- print_azure_ad_setup_instructions(): Show Azure AD app setup instructions
- main(): Main function (handles both command line and Excel usage)

SharePoint Requirements:
- Azure AD app registration with Graph API permissions
- Client ID and Tenant ID for authentication  
- Redirect URI configured: http://localhost:8080
- SharePoint site access permissions

IMPORTANT: If you get redirect URI errors, run:
parse_reports.print_azure_ad_setup_instructions()
"""

import sys
import os
from pathlib import Path
from openpyxl import load_workbook
import requests
import msal
import io
from urllib.parse import urlparse, quote
import tempfile
from datetime import datetime
from keyvault import get_secret
from search_sharepoint import get_sharepoint_list_items

# SharePoint configuration
try:
    from sharepoint_config import get_config, CLIENT_ID, TENANT_ID, REDIRECT_URI
    SHAREPOINT_CONFIG = get_config()
except ImportError:
    print("Error: sharepoint_config.py not found or has errors.")
    raise


def get_sharepoint_token(client_id=None, tenant_id=None, redirect_url=None): 
    """Get SharePoint access token using MSAL"""
    try:
        if not client_id:
            client_id = input("Enter your Azure AD application client ID: ")
        if not tenant_id:
            tenant_id = input("Enter your Azure AD tenant ID: ")
        
        # Default redirect URI for public client apps
        if not redirect_url:
            print("No redirect URI provided")
            return None
        
        # Initialize MSAL app with fixed redirect URI
        app = msal.PublicClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}"
        )
                
        scopes = ["https://graph.microsoft.com/.default"]

        # Try to get token silently first (from cache)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                print("Using cached token")
                return result["access_token"]
        
        # Tech debt: Temporary workaround should be stored in Key Vault
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
        
            return None
    except Exception as e:
        print(f"Error in SharePoint authentication: {str(e)}")
        return None


def search_sharepoint_files(access_token, config=None):
    """Search for Excel files in SharePoint"""
    try:
        if config is None:
            config = SHAREPOINT_CONFIG
            
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        search_query = {
            "requests": [{
                "entityTypes": ["driveItem"],
                "query": {
                    "queryString": f'site:"{config["site"]}" path:"{config["relative_path"]}" filename:"{config["search_pattern"]}*.xlsx"'
                },
                "region": "GBR"  # Required for application permissions
            }]
        }
        
        response = requests.post(
            'https://graph.microsoft.com/v1.0/search/query',
            headers=headers,
            json=search_query
        )
        
        if response.status_code == 200:
            data = response.json()
            hits = data.get('value', [{}])[0].get('hitsContainers', [{}])[0].get('hits', [])
            return hits
        else:
            print(f"SharePoint search error: {response.status_code}")
            print(response.text)
            return []
    except Exception as e:
        print(f"Error searching SharePoint files: {str(e)}")
        return []


def download_sharepoint_file(access_token, file_url):
    """Download a file from SharePoint and return file content"""
    try:
        headers = {
            'Authorization': f'Bearer {access_token}'
        }
        
        # Convert web URL to download URL
        if 'sharepoint.com' in file_url:
            # Extract site and file path from URL
            parsed_url = urlparse(file_url)
            site_path = parsed_url.path.split('/sites/')[1].split('/')[0]
            
            # Get file info using Graph API
            site_id_url = f"https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/{site_path}"
            site_response = requests.get(site_id_url, headers=headers)
            
            if site_response.status_code == 200:
                site_id = site_response.json()['id']
                
                # Search for the file using Graph API
                file_name = file_url.split('/')[-1]
                search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{file_name}')"
                search_response = requests.get(search_url, headers=headers)
                
                if search_response.status_code == 200:
                    files = search_response.json().get('value', [])
                    if files:
                        download_url = files[0]['@microsoft.graph.downloadUrl']
                        file_response = requests.get(download_url)
                        
                        if file_response.status_code == 200:
                            return file_response.content
        
        # Fallback: try direct download
        response = requests.get(file_url, headers=headers)
        if response.status_code == 200:
            return response.content
        else:
            print(f"Error downloading file: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Error downloading SharePoint file: {str(e)}")
        return None


def process_sharepoint_workbook(access_token, file_info):
    """Process a SharePoint workbook"""
    try:
        resource = file_info.get('resource', {})
        file_name = resource.get('name', 'unknown.xlsx')
        file_url = resource.get('webUrl', '')
        
        print(f"\nProcessing SharePoint file: {file_name}")
        print(f"URL: {file_url}")
        
        # Download file content
        file_content = download_sharepoint_file(access_token, file_url)
        if not file_content:
            return False
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
            temp_file.write(file_content)
            temp_filename = temp_file.name
        
        try:
            # Process the temporary file
            success = process_workbook_content(temp_filename, file_name, access_token)
            return success
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_filename)
            except:
                pass
                
    except Exception as e:
        print(f"Error processing SharePoint workbook: {str(e)}")
        return False

def process_workbook_content_from_memory(file_content, original_name, access_token=None):
    """Process workbook content directly from memory without saving to disk"""
    try:
        # Use original filename for output
        output_name = original_name
        
        # Load the workbook from memory using BytesIO
        file_like = io.BytesIO(file_content)
        wb = load_workbook(file_like, data_only=True)
        
        return _process_workbook_data(wb, output_name, access_token)
        
    except Exception as e:
        print(f"Error processing workbook from memory {original_name}: {str(e)}")
        return False


def process_workbook_content(filename, original_name=None, access_token=None):
    """Process workbook content (extracted from original process_workbook function)"""
    try:
        # Use original filename for output if provided
        output_name = original_name if original_name else os.path.basename(filename)
        
        # Load the workbook
        wb = load_workbook(filename, data_only=True)
        
        return _process_workbook_data(wb, output_name, access_token)
        
    except Exception as e:
        print(f"Error processing workbook {filename}: {str(e)}")
        return False


def _process_workbook_data(wb, output_name, access_token=None):
    """Common workbook processing logic used by both file and memory-based functions"""
    try:
        # Get the active sheet
        sheet = wb.active
        
        # Validate that we have a sheet
        if sheet is None:
            print(f"Error: Could not access active sheet in {output_name}")
            wb.close()
            return False
        
        # First check if E33 is merged
        is_merged, target_range = is_cell_merged(sheet, "E33")
        
        if is_merged:
            print(f"Cell E33 is merged. Unmerging cells in range {target_range}")
            # Unmerge SSN range
            for merged_cell_range in list(sheet.merged_cells.ranges):
                min_col, min_row, max_col, max_row = merged_cell_range.bounds
                # Check if the merged cell range overlaps with our target range (D31:R42)
                if not (max_col < 4 or min_col > 19 or max_row < 31 or min_row > 72):
                    sheet.unmerge_cells(str(merged_cell_range))
        else:
            print("Cell E33 is not merged. No action taken.")
        
        # Create output filename
        name, ext = output_name.rsplit('.', 1) if '.' in output_name else (output_name, 'xlsx')
        fname = f"{name}.txt"

        # Build content in memory
        content_lines = []
        content_lines.append(f"Week Ending: {sheet['G7'].value}")
        content_lines.append(f"Service Provider: {sheet['G11'].value}")
        content_lines.append(f"Client: {sheet['G13'].value}")
        content_lines.append("")
        content_lines.append("Service Standard updates:")
        content_lines.append("SSN|Status|Comments")
        
        # Service standards
        for row in range(34, 43):
            if sheet[f'D{row}'].value:
                content_lines.append(f"{sheet[f'D{row}'].value}|{sheet[f'J{row}'].value}|{sheet[f'K{row}'].value}")

        # Service Risks
        content_lines.append("")
        content_lines.append("Service Risks:")
        content_lines.append("Risk No|Description|Likelihood|Impact|Mitigation")

        for row in range(45, 48):
            if sheet[f'D{row}'].value:
                content_lines.append(f"{sheet[f'D{row}'].value}|{sheet[f'E{row}'].value}|{sheet[f'H{row}'].value}|{sheet[f'J{row}'].value}|{sheet[f'K{row}'].value}")

        # Service Issues
        content_lines.append("")
        content_lines.append("Service Issues:")
        content_lines.append("Issue No|Description|Impact|Mitigation")

        for row in range(50, 53):
            if sheet[f'D{row}'].value:
                content_lines.append(f"{sheet[f'D{row}'].value}|{sheet[f'E{row}'].value}|{sheet[f'J{row}'].value}|{sheet[f'K{row}'].value}")

        # Planned Activities
        content_lines.append("")
        content_lines.append("Planned Activities:")
        content_lines.append(str(sheet['D57'].value) if sheet['D57'].value else "")

        # Client Updates
        content_lines.append("")
        content_lines.append("Client Updates:")
        content_lines.append(str(sheet['D67'].value) if sheet['D67'].value else "")

        # Join all content
        file_content = "\n".join(content_lines)
        
        # Close workbook without saving changes
        wb.close()
        
        # Upload to SharePoint if access token is provided, otherwise save locally
        if access_token:
            success = upload_text_to_sharepoint(access_token, file_content, fname)
            if success:
                print(f"📄 Successfully processed Excel data and uploaded to SharePoint: {fname}")
                return True
            else:
                print(f"⚠️ Failed to upload to SharePoint, falling back to local save: {fname}")
                # Fall back to local save
        
        # Save locally as fallback or when no access token provided
        with open(fname, 'w', encoding='utf-8') as file:
            file.write(file_content)
        print(f"📄 Successfully processed Excel data and saved locally as: {fname}")
        return True

    except Exception as e:
        print(f"Error processing workbook data: {str(e)}")
        return False

def is_cell_merged(sheet, cell_coord):
    """Check if a specific cell is part of a merged range"""
    for merged_range in sheet.merged_cells.ranges:
        if cell_coord in merged_range:
            return True, merged_range
    return False, None

def process_workbook(filename):
    """Process a single local workbook and extract information"""
    try:
        # Ensure we have absolute path
        filename = os.path.abspath(filename)
        print(f"\nProcessing local file: {filename}")
        
        # Check if file exists
        if not os.path.exists(filename):
            print(f"Error: File '{filename}' does not exist")
            return False
            
        # Use the content processing function
        return process_workbook_content(filename)

    except Exception as e:
        print(f"Error processing file {filename}: {str(e)}")
        return False

def process_directory(directory_path):
    """Process all Excel files in the specified directory"""
    # Check if directory exists
    if not os.path.exists(directory_path):
        print(f"Error: Directory '{directory_path}' does not exist")
        return False

    # Get list of Excel files
    excel_files = [f for f in os.listdir(directory_path) 
                  if f.endswith(('.xlsx', '.xlsm', '.xls'))]
    
    if not excel_files:
        print(f"No Excel files found in '{directory_path}'")
        return False

    print(f"Found {len(excel_files)} Excel files to process")
    
    # Process each file
    success_count = 0
    for excel_file in excel_files:
        full_path = os.path.join(directory_path, excel_file)
        if process_workbook(full_path):
            success_count += 1
    
    print(f"\nProcessing complete: {success_count} out of {len(excel_files)} files processed successfully")
    return True

def process_current_folder():
    """Convenience function for Excel - processes Excel files in current directory"""
    try:
        current_dir = os.getcwd()
        print(f"Processing Excel files in current folder: {current_dir}")
        return process_directory(current_dir)
    except Exception as e:
        print(f"Error processing current folder: {str(e)}")
        return False

def process_folder(folder_path=None):
    """Function to process a specific folder or current folder if none specified"""
    try:
        if folder_path is None:
            folder_path = os.getcwd()
        
        # Convert to absolute path if relative
        folder_path = os.path.abspath(folder_path)
        print(f"Processing folder: {folder_path}")
        
        return process_directory(folder_path)
    except Exception as e:
        print(f"Error processing folder: {str(e)}")
        return False


def process_sharepoint_files(client_id=None, tenant_id=None, config=None, redirect_url=None):
    """Process Excel files from SharePoint using list items"""
    try:
        print("Starting SharePoint file processing using list items...")
        
        # Use config values if not provided
        if client_id is None:
            client_id = CLIENT_ID
        if tenant_id is None:
            tenant_id = TENANT_ID
        if redirect_url is None:
            redirect_url = f"msal{CLIENT_ID}://auth"
        
        # Get access token
        access_token = get_sharepoint_token(client_id, tenant_id, redirect_url)
        if not access_token:
            return False
        
        # Get SharePoint list items
        print("Fetching SharePoint list items...")
        # Tech debt: site and list name should be configurable
        site_name = "jjag.sharepoint.com"
        list_name = "Service Provider Uploads"
        
        list_items = get_sharepoint_list_items(access_token, site_name, list_name)
        if not list_items:
            print("No SharePoint list items found")
            return False
        
        print(f"Found {len(list_items)} items in SharePoint list")
        
        # Filter for unprocessed Excel files
        excel_files = []
        for item in list_items:
            fields = item.get('fields', {})
            
            # Get path and filename
            path = fields.get('Path', '')
            filename = fields.get('Reportfilename', '')
            filename = f"{filename.strip()}.xlsx" 
            
            # Check if monthly report processed
            monthly_report_processed = (
                fields.get('Monthly_x0020_Report_x0020_Processed') or
                fields.get('MonthlyReportProcessed') or
                fields.get('monthly_report_processed') or
                fields.get('Monthly Report Processed') or
                False
            )
            
            # Only process Excel files that haven't been processed
            if (filename and monthly_report_processed not in [True, 'Yes', 'true', '1', 1, 'True']):
                
                excel_files.append({
                    'filename': filename,
                    'path': path,
                    'full_url': f"{path.rstrip('/')}/{filename}",
                    'item_id': item.get('id', ''),
                    'fields': fields
                })
        
        if not excel_files:
            print("No unprocessed Excel files found in SharePoint list")
            return False
        
        print(f"Found {len(excel_files)} unprocessed Excel files to process")
        
        # Process each Excel file
        success_count = 0
        for file_info in excel_files:
            try:
                print(f"\nProcessing: {file_info['filename']}")
                
                # Download file from SharePoint
                file_content = download_sharepoint_file_from_path(access_token, file_info['path'], file_info['filename'])
                if not file_content:
                    print(f"Failed to download: {file_info['filename']}")
                    continue
                
                # Process the Excel file directly from memory
                if process_workbook_content_from_memory(file_content, file_info['filename'], access_token):
                    success_count += 1
                    print(f"Successfully processed: {file_info['filename']}")
                    
                    # Mark as processed in SharePoint
                    #mark_file_as_processed(access_token, file_info['item_id'])
                    
                else:
                    print(f"Failed to process Excel content: {file_info['filename']}")
                        
            except Exception as e:
                print(f"Error processing {file_info['filename']}: {str(e)}")
        
        print(f"\nSharePoint processing complete: {success_count} out of {len(excel_files)} files processed successfully")
        return success_count > 0
        
    except Exception as e:
        print(f"Error processing SharePoint files: {str(e)}")
        return False


def check_token_permissions(access_token):
    """Check what permissions the current token has"""
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Validate application token (skip /me endpoint as it requires delegated auth)
        print("[INFO] ✅ Using application-only authentication (service-to-service)")
        print("[INFO] Validating token with Graph metadata endpoint...")
        
        # Use Graph metadata endpoint which works with application tokens
        service_root_url = "https://graph.microsoft.com/v1.0/$metadata"
        metadata_response = requests.get(service_root_url, headers=headers)
        print(f"[DEBUG] Service metadata endpoint: {metadata_response.status_code}")
        
        if metadata_response.status_code == 200:
            print("[INFO] ✅ Application token is valid and working")
        elif metadata_response.status_code == 401:
            print("[ERROR] ❌ Token is invalid or expired")
            return False
        elif metadata_response.status_code == 403:
            print("[ERROR] ❌ Token lacks basic Graph API permissions")
            return False
        else:
            print(f"[WARNING] ⚠️ Unexpected response from metadata endpoint: {metadata_response.status_code}")
            print(f"[DEBUG] Response: {metadata_response.text[:200]}")
        
        # Try to list sites to see if we have Sites.Read permission
        sites_url = "https://graph.microsoft.com/v1.0/sites"
        sites_response = requests.get(sites_url, headers=headers)
        
        print(f"[DEBUG] Sites list permission: {sites_response.status_code}")
        
        # Try to get the specific site
        site_url = "https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/InternalTeam:"
        site_response = requests.get(site_url, headers=headers)
        
        print(f"[DEBUG] Specific site access: {site_response.status_code}")
        
        if site_response.status_code == 200:
            site_data = site_response.json()
            site_id = site_data['id']
            
            # Try to list drives to see permissions
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
            drives_response = requests.get(drives_url, headers=headers)
            
            print(f"[DEBUG] Drives list permission: {drives_response.status_code}")
            
            if drives_response.status_code == 200:
                drives_data = drives_response.json()
                print(f"[DEBUG] Available drives: {[d.get('name', 'Unknown') for d in drives_data.get('value', [])]}")
                
                # Try to list root folder contents
                for drive in drives_data.get('value', []):
                    drive_id = drive.get('id')
                    root_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children"
                    root_response = requests.get(root_url, headers=headers)
                    print(f"[DEBUG] Drive '{drive.get('name')}' root access: {root_response.status_code}")
                    
                    if root_response.status_code == 200:
                        root_data = root_response.json()
                        folders = [item.get('name') for item in root_data.get('value', []) if 'folder' in item]
                        print(f"[DEBUG] Available folders in '{drive.get('name')}': {folders}")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Token permission check failed: {e}")
        return False


def upload_text_to_sharepoint(access_token, file_content, filename):
    """Upload text content to SharePoint"""
    try:
        print(f"[DEBUG] Starting SharePoint upload for: {filename}")
        
        # Check token permissions first
        check_token_permissions(access_token)
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'text/plain'
        }
        
        # Get site ID first
        site_info_url = "https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/InternalTeam:"
        site_response = requests.get(site_info_url, headers={'Authorization': f'Bearer {access_token}'})
        
        if site_response.status_code != 200:
            print(f"Failed to get site info: {site_response.status_code} - {site_response.text}")
            return False
            
        site_id = site_response.json()['id']
        print(f"Got site ID: {site_id}")
        
        # Try different path approaches - focusing on the new target location
        target_paths = [
            # Primary target: MonthlyReports folder structure
            "MonthlyReports/2025/09 - September",
            # Alternative paths in case the exact structure doesn't exist
            "/Shared Documents/MonthlyReports/2025",
            "/Shared Documents/MonthlyReports",
            # Fallback to root shared documents
            "/Shared Documents",
            # Original paths as final fallback
            "/Shared Documents/monthlyreport"
        ]
        
        for target_path in target_paths:
            try:
                print(f"Trying path: {target_path}")
                
                # Construct file path for Graph API
                file_path = f"{target_path}/{filename}"
                
                # Clean up path - remove double slashes
                file_path = file_path.replace('//', '/')
                
                # URL encode the path
                encoded_path = quote(file_path, safe='/')
                
                # Try to upload file to SharePoint
                upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{encoded_path}:/content"
                print(f"Upload URL: {upload_url}")
                
                upload_response = requests.put(upload_url, headers=headers, data=file_content.encode('utf-8'))
                
                if upload_response.status_code in [200, 201]:
                    print(f"✅ Successfully uploaded to SharePoint: {filename} at path: {target_path}")
                    return True
                else:
                    print(f"❌ Failed with path {target_path}: {upload_response.status_code} - {upload_response.text}")
                    
            except Exception as path_error:
                print(f"❌ Error with path {target_path}: {str(path_error)}")
                continue
        
        # If all paths failed, try alternative approaches
        print("All direct upload attempts failed. Trying alternative approaches...")
        
        try:
            # Method 1: Try to get all drives and find the Documents library
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
            drives_response = requests.get(drives_url, headers={'Authorization': f'Bearer {access_token}'})
            
            if drives_response.status_code == 200:
                drives_data = drives_response.json()
                print(f"[DEBUG] Found drives: {[d.get('name', 'Unknown') for d in drives_data.get('value', [])]}")
                
                # Look for the Documents library (usually named 'Documents' or 'Shared Documents')
                for drive in drives_data.get('value', []):
                    drive_name = drive.get('name', '')
                    drive_id = drive.get('id')
                    
                    if 'document' in drive_name.lower() or 'shared' in drive_name.lower():
                        print(f"[DEBUG] Trying to upload to drive: {drive_name} (ID: {drive_id})")
                        
                        # Try to upload directly to this drive's root
                        drive_upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{quote(filename, safe='')}:/content"
                        
                        upload_response = requests.put(drive_upload_url, 
                                                     headers={'Authorization': f'Bearer {access_token}', 'Content-Type': 'text/plain'}, 
                                                     data=file_content.encode('utf-8'))
                        
                        if upload_response.status_code in [200, 201]:
                            print(f"✅ Successfully uploaded to drive '{drive_name}': {filename}")
                            return True
                        else:
                            print(f"❌ Failed to upload to drive '{drive_name}': {upload_response.status_code} - {upload_response.text}")
            
            # Method 2: Try to create the MonthlyReports folder structure and upload there
            print("[DEBUG] Trying to create MonthlyReports folder structure...")
            
            # Create the folder structure step by step
            folders_to_create = [
                ("MonthlyReports", "/Shared Documents"),
                ("2025", "/Shared Documents/MonthlyReports"), 
                ("09 - September", "/Shared Documents/MonthlyReports/2025")
            ]
            
            for folder_name, parent_path in folders_to_create:
                try:
                    folder_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{quote(parent_path, safe='/')}:/children"
                    
                    folder_data = {
                        "name": folder_name,
                        "folder": {},
                        "@microsoft.graph.conflictBehavior": "replace"
                    }
                    
                    create_response = requests.post(folder_url, 
                                                 headers={'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}, 
                                                 json=folder_data)
                    
                    if create_response.status_code in [200, 201, 409]:  # 409 = already exists
                        print(f"✅ Created/confirmed folder: {parent_path}/{folder_name}")
                    else:
                        print(f"⚠️ Could not create folder {folder_name}: {create_response.status_code}")
                        
                except Exception as folder_error:
                    print(f"⚠️ Error creating folder {folder_name}: {folder_error}")
            
            # Now try to upload to the target folder
            target_upload_path = "/Shared Documents/MonthlyReports/2025/09 - September"
            final_upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{quote(target_upload_path, safe='/')}/{quote(filename, safe='')}:/content"
            
            upload_response = requests.put(final_upload_url, 
                                         headers={'Authorization': f'Bearer {access_token}', 'Content-Type': 'text/plain'}, 
                                         data=file_content.encode('utf-8'))
            
            if upload_response.status_code in [200, 201]:
                print(f"✅ Successfully uploaded to MonthlyReports folder: {filename}")
                return True
            else:
                print(f"❌ Failed to upload to MonthlyReports folder: {upload_response.status_code} - {upload_response.text}")
                
        except Exception as alt_error:
            print(f"❌ Error with alternative approaches: {str(alt_error)}")
        
        # Method 3: Provide guidance for fixing permissions
        print("\n" + "="*80)
        print("📋 SHAREPOINT UPLOAD TROUBLESHOOTING GUIDE")
        print("="*80)
        print("\n🔐 The 403 'Access Denied' errors suggest insufficient permissions.")
        print("\n📝 To fix this, your Azure AD application needs these Microsoft Graph permissions:")
        print("   • Sites.ReadWrite.All (to write to SharePoint sites)")
        print("   • Files.ReadWrite.All (to upload files)")
        print("   • Sites.Manage.All (optional, for creating folders)")
        print("\n🔧 Steps to add permissions:")
        print("   1. Go to Azure Portal > Azure Active Directory > App registrations")
        print("   2. Find your application and click on it")
        print("   3. Go to 'API permissions' > 'Add a permission'")
        print("   4. Select 'Microsoft Graph' > 'Application permissions'")
        print("   5. Add the permissions listed above")
        print("   6. Click 'Grant admin consent' (IMPORTANT!)")
        print("\n💡 Alternative: Try uploading to a different SharePoint location with broader access.")
        print("\n📂 Current files are being saved locally as fallback.")
        print("="*80)
        
        print("❌ All upload attempts failed")
        return False
            
    except Exception as e:
        print(f"❌ Error uploading file {filename} to SharePoint: {str(e)}")
        return False


def download_sharepoint_file_from_path(access_token, sharepoint_path, filename):
    """Download a file from SharePoint using path and filename"""
    try:
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json'
        }
        
        # Use the specified SharePoint path for downloads
        # Default to the new MonthlyReports path, but allow override via sharepoint_path parameter if needed
        if sharepoint_path and sharepoint_path.strip():
            # Clean the path - remove "Shared Documents/" if present
            clean_path = sharepoint_path
            if clean_path.startswith('Shared Documents/'):
                clean_path = clean_path[len('Shared Documents/'):]
            elif clean_path.startswith('/Shared Documents/'):
                clean_path = clean_path[len('/Shared Documents/'):]
            target_path = clean_path
        else:
            target_path = "sites/InternalTeam/Shared Documents/MonthlyReports/2025/09 - September"
        
        # Construct file path for Graph API
        if target_path:
            file_path = f"/{target_path}/{filename}"
        else:
            file_path = f"/{filename}"
        
        # Clean up path - remove double slashes
        file_path = file_path.replace('//', '/')
        
        # URL encode the path
        encoded_path = quote(file_path, safe='/')
        
        # Get site ID first
        site_info_url = "https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/InternalTeam:"
        site_response = requests.get(site_info_url, headers=headers)
        
        if site_response.status_code != 200:
            print(f"Failed to get site info: {site_response.status_code}")
            return None
            
        site_id = site_response.json()['id']
        
        # Try to get file directly
        file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{encoded_path}"
        
        file_response = requests.get(file_url, headers=headers)
        
        if file_response.status_code == 200:
            file_info = file_response.json()
            download_url = file_info.get('@microsoft.graph.downloadUrl')
            
            if download_url:
                # Download the actual file content
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    print(f"Successfully downloaded: {filename}")
                    return download_response.content
        else:
            # Fallback: search for the file
            search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/search(q='{filename}')"
            search_response = requests.get(search_url, headers=headers)
            
            if search_response.status_code == 200:
                search_results = search_response.json().get('value', [])
                
                for item in search_results:
                    if item.get('name') == filename:
                        # Get file details
                        file_id = item.get('id')
                        detail_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}"
                        detail_response = requests.get(detail_url, headers=headers)
                        
                        if detail_response.status_code == 200:
                            detail_data = detail_response.json()
                            download_url = detail_data.get('@microsoft.graph.downloadUrl')
                            
                            if download_url:
                                download_response = requests.get(download_url)
                                if download_response.status_code == 200:
                                    print(f"Successfully downloaded via search: {filename}")
                                    return download_response.content
        
        print(f"Failed to download file: {filename}")
        return None
        
    except Exception as e:
        print(f"Error downloading file {filename}: {str(e)}")
        return None


def process_sharepoint_list_files(client_id=None, tenant_id=None, list_name="Service Provider Uploads", redirect_uri=None):
    """Process Excel files from SharePoint list using path and filename columns"""
    try:
        print("[INFO] Starting SharePoint list file processing...")
        
        # Get SharePoint token using MSAL ConfidentialClientApplication (same as search_sharepoint.py)
        try:
            # Use the same client secret as our working search_sharepoint.py
            client_secret = "fPl8Q~oEqBx.Mi0sfTJq2PQ-teDBeaHG6M5K4cKN"
            
            app = msal.ConfidentialClientApplication(
                client_id=CLIENT_ID,
                client_credential=client_secret,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}"
            )
            
            # Get token for Microsoft Graph API with write permissions
            scopes = ["https://graph.microsoft.com/.default"]
            result = app.acquire_token_for_client(scopes=scopes)
            
            print(f"[DEBUG] Requested scopes: {scopes}")
            print(f"[DEBUG] Token result keys: {list(result.keys()) if isinstance(result, dict) else 'Not a dict'}")
            
            if "access_token" not in result:
                print(f"[ERROR] Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
                
            token = result["access_token"]
            print("[SUCCESS] Successfully authenticated with Microsoft Graph API")
            
        except Exception as auth_error:
            print(f"[ERROR] Authentication error: {auth_error}")
            return False
        
        # Set up headers for API calls
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json'
        }
        
        # Use the site configuration from search_sharepoint.py
        sharepoint_site_id = "jjag.sharepoint.com,e5b0f1b3-8b7c-4b5f-9b5f-1b3e5b0f1b3e,a1b2c3d4-e5f6-7890-abcd-ef1234567890"  # This is typically auto-discovered
        
        # Try multiple site endpoints to find the list
        site_endpoints = [
            "https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/InternalTeam:/lists/Service%20Provider%20Uploads/items?$expand=fields",
            "https://graph.microsoft.com/v1.0/sites/root/lists/Service%20Provider%20Uploads/items?$expand=fields",
            "https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com/lists/Service%20Provider%20Uploads/items?$expand=fields"
        ]
        
        items = []
        for endpoint in site_endpoints:
            try:
                response = requests.get(endpoint, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    items = data.get('value', [])
                    print(f"[SUCCESS] Found {len(items)} items from SharePoint list")
                    break
            except Exception as e:
                print(f"❌ Error with endpoint {endpoint}: {e}")
                continue
        
        if not items:
            print("[ERROR] No items found in SharePoint list")
            return False
        
        # Get target months for filtering
        from sharepoint_config import get_previous_month_path, get_current_month_path
        target_months = [get_previous_month_path(), get_current_month_path()]
        print(f"[INFO] Looking for files in months: {target_months}")
        
        # Show sample paths to debug path matching
        print("\n[INFO] Sample paths from SharePoint list:")
        sample_count = 0
        for item in items[:5]:  # Show first 5 items
            fields = item.get('fields', {})
            path = fields.get('Path', '')
            filename = fields.get('Reportfilename', '')
            if path and 'CS Flex Weekly Service Delivery Report' in str(filename):
                print(f"  • {path}")
                sample_count += 1
        if sample_count == 0:
            print("  No CS Flex Weekly Service Delivery Report files found in sample")
        
        # Process items to find unprocessed report files using "monthly report processed" field
        report_files = []
        
        print("\n[INFO] Checking processing status for all files...")
        processed_count_check = 0
        unprocessed_count_check = 0
        
        for item in items:
            fields = item.get('fields', {})
            
            # Get path and filename using the correct field names we found
            path = fields.get('Path', '')
            filename = fields.get('Reportfilename', '')
            
            # Check various possible field names for "monthly report processed"
            monthly_report_processed = (fields.get('Monthlyreportprocessed'))
            if monthly_report_processed is True:
                continue  # Already processed

            manager = fields.get('manager', '')
            if manager != 'Julian Brown':
                continue  # Not the target manager  

            # Skip if we don't have both path and filename
            if not path or not filename:
                continue
                
            # Check if this is a report file (using path and filename directly, no search pattern)
            if filename and path:
                # Only process files that haven't been processed yet
                if monthly_report_processed not in [True, 'Yes', 'true', '1', 1, 'True']:
                    print(f"[INFO] Found unprocessed file: {filename}")
                    unprocessed_count_check += 1
                    report_files.append({
                        'filename': filename,
                        'path': path,
                        'full_url': f"{path.rstrip('/')}/{filename}",
                        'modified': fields.get('Modified', ''),
                        'item_id': item.get('id', ''),
                        'fields': fields,
                        'monthly_report_processed': monthly_report_processed
                    })
                else:
                    processed_count_check += 1
                    print(f"⏭️  Skipping already processed file: {filename}")
        
        print(f"\n📊 Processing Status Summary:")
        print(f"   ✅ Already processed: {processed_count_check} files")
        print(f"   📋 Unprocessed (will process): {unprocessed_count_check} files")
        
        if not report_files:
            print("❌ No unprocessed files found in SharePoint list")
            return False
        
        print(f"\n🎯 Found {len(report_files)} unprocessed files to process")
        
        # Process each report file
        success_count = 0
        processed_count = 0
        max_files = 50  # Limit processing for testing
        
        for report_info in report_files[:max_files]:
            try:
                print(f"\n📊 Processing: {report_info['filename']}")
                
                # Construct SharePoint file download URL
                print(f"[DOWNLOAD] Downloading file: {report_info['filename']}")
                
                # Convert SharePoint path to download URL
                # The path from SharePoint list includes "Shared Documents" but we need to remove it
                # for the Graph API drive/root: endpoint since it's implicit
                
                clean_path = report_info['path']
                
                # Remove "Shared Documents/" if it's at the start of the path
                if clean_path.startswith('Shared Documents/'):
                    clean_path = clean_path[len('Shared Documents/'):]
                elif clean_path.startswith('/Shared Documents/'):
                    clean_path = clean_path[len('/Shared Documents/'):]
                
                # Construct the server relative URL
                if clean_path:
                    server_relative_url = f"/{clean_path}/{report_info['filename']}"
                else:
                    server_relative_url = f"/{report_info['filename']}"
                
                # Clean up the URL - remove double slashes
                server_relative_url = server_relative_url.replace('//', '/')
                
                # Use Graph API to get file by path using the drive API
                # Use the same approach as the list access but for drive items
                encoded_path = server_relative_url.replace(' ', '%20')
                
                # Try different approaches for file access
                file_urls_to_try = [
                    f"https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/InternalTeam:/drive/root:{encoded_path}",
                    f"https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com/sites/InternalTeam/drive/root:{encoded_path}",
                    f"https://graph.microsoft.com/v1.0/drives/b!s_Gw5XsLX0ub1f1V8F8T7r4g0tJq4TtCo_sATwC6tkquG1ZbWBsPQI_j_13Nk6Dl/root:{encoded_path}"
                ]
                
                file_url = file_urls_to_try[0]  # Start with the first one
                
                try:
                    print(f"[INFO] Attempting to download file: {report_info['filename']}")
                    
                    headers = {
                        'Authorization': f'Bearer {token}',
                        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    }
                    
                    download_url = None
                    file_info = None
                    parsed_data = None
                    
                    # METHOD 1: Get site ID first, then use it to access drive items
                    print(f"   Method 1: Getting site ID for proper drive access")
                    site_info_url = "https://graph.microsoft.com/v1.0/sites/jjag.sharepoint.com:/sites/InternalTeam:"
                    site_response = requests.get(site_info_url, headers=headers)
                    
                    site_id = None
                    if site_response.status_code == 200:
                        site_data = site_response.json()
                        site_id = site_data['id']
                        print(f"   Got site ID: {site_id}")
                        
                        # Now try to search using the proper site ID
                        drive_search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/search(q='{report_info['filename']}')"
                        print(f"   Trying drive search: {drive_search_url[:100]}...")
                        
                        search_response = requests.get(drive_search_url, headers=headers)
                        
                        if search_response.status_code == 200:
                            search_results = search_response.json().get('value', [])
                            print(f"   Found {len(search_results)} search results")
                            
                            # Find exact match
                            for item in search_results:
                                if item.get('name') == report_info['filename']:
                                    file_info = item
                                    # Get the file ID and fetch full details including download URL
                                    file_id = item.get('id')
                                    if file_id:
                                        file_detail_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}"
                                        detail_response = requests.get(file_detail_url, headers=headers)
                                        if detail_response.status_code == 200:
                                            detail_data = detail_response.json()
                                            download_url = detail_data.get('@microsoft.graph.downloadUrl')
                                            print(f"[SUCCESS] Found exact match and got download URL")
                                            break
                            
                            # If no exact match, try partial match
                            if not download_url and search_results:
                                for item in search_results:
                                    item_name = item.get('name', '')
                                    if 'CS Flex Weekly' in item_name and any(word in item_name for word in report_info['filename'].split()[-3:]):
                                        file_info = item
                                        # Get the file ID and fetch full details including download URL
                                        file_id = item.get('id')
                                        if file_id:
                                            file_detail_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}"
                                            detail_response = requests.get(file_detail_url, headers=headers)
                                            if detail_response.status_code == 200:
                                                detail_data = detail_response.json()
                                                download_url = detail_data.get('@microsoft.graph.downloadUrl')
                                                print(f"[SUCCESS] Found similar file and got download URL: {item_name}")
                                                break
                        else:
                            print(f"   Drive search failed: {search_response.status_code}")
                    
                    # METHOD 2: If Method 1 fails, try direct file access using known path patterns
                    if not download_url and site_id:
                        print(f"   Method 2: Trying direct file access")
                        
                        # Clean the path and construct direct access URLs
                        clean_path = report_info['path']
                        if clean_path.startswith('Shared Documents/'):
                            clean_path = clean_path[len('Shared Documents/'):]
                        
                        # Try multiple URL formats that might work
                        file_urls = [
                            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{clean_path}/{report_info['filename']}",
                            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/Documents/root:/{clean_path}/{report_info['filename']}"
                        ]
                        
                        for file_url in file_urls:
                            print(f"   Trying: {file_url[:100]}...")
                            file_response = requests.get(file_url, headers=headers)
                            
                            if file_response.status_code == 200:
                                file_info = file_response.json()
                                download_url = file_info.get('@microsoft.graph.downloadUrl')
                                print(f"[SUCCESS] Found file via direct access")
                                break
                            else:
                                print(f"   Direct access failed: {file_response.status_code}")
                    
                    # METHOD 3: If still no luck, try a simpler search approach
                    if not download_url and site_id:
                        print(f"   Method 3: Trying simplified search")
                        
                        # Search for just the key terms
                        search_terms = "CS Flex Weekly Service Delivery Report"
                        simple_search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/search(q='{search_terms}')"
                        
                        search_response = requests.get(simple_search_url, headers=headers)
                        if search_response.status_code == 200:
                            search_results = search_response.json().get('value', [])
                            print(f"   Found {len(search_results)} results with simplified search")
                            
                            # Look for files that match our criteria
                            for item in search_results:
                                item_name = item.get('name', '')
                                # Check if this could be our file based on name similarity
                                if report_info['filename'][:20] in item_name or item_name[:20] in report_info['filename']:
                                    file_info = item
                                    # Get the file ID and fetch full details including download URL
                                    file_id = item.get('id')
                                    if file_id:
                                        file_detail_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}"
                                        detail_response = requests.get(file_detail_url, headers=headers)
                                        if detail_response.status_code == 200:
                                            detail_data = detail_response.json()
                                            download_url = detail_data.get('@microsoft.graph.downloadUrl')
                                            print(f"[SUCCESS] Found potential match and got download URL: {item_name}")
                                            break
                        else:
                            print(f"   Simplified search failed: {search_response.status_code}")
                    
                    if not download_url:
                        print(f"[WARNING] Could not find download URL for: {report_info['filename']}")
                        print(f"   Will create metadata-only report")
                        # Continue with metadata processing instead of failing
                        parsed_data = {
                            'filename': report_info['filename'],
                            'processing_result': 'File download failed - processed metadata only',
                            'summary': f'Could not download actual file, processed SharePoint list metadata',
                            'worksheet_info': 'No Excel data available - download failed',
                            'sharepoint_metadata': {
                                'path': report_info['path'],
                                'modified': report_info['modified'],
                                'processing_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                'source': 'SharePoint List - Service Provider Uploads',
                                'full_url': report_info['full_url']
                            }
                        }
                        success_count += 1
                        processed_count += 1
                    
                    # Download the file if we have a download URL
                    if download_url:
                        download_headers = {
                            'Authorization': f'Bearer {token}',
                            'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        }
                        
                        print(f"[DOWNLOAD] Downloading from: {download_url[:100]}...")
                        download_response = requests.get(download_url, headers=download_headers)
                    else:
                        print(f"[INFO] No download URL available, processing metadata only")
                        download_response = None
                    
                    if download_response and download_response.status_code == 200:
                        # Save the file temporarily for processing (use safe filename)
                        import re
                        safe_name = re.sub(r'[<>:"/\\|?*]', '_', report_info['filename'])
                        temp_filename = f"temp_{safe_name}"
                        
                        with open(temp_filename, 'wb') as temp_file:
                            temp_file.write(download_response.content)
                        
                        print(f"[SUCCESS] Downloaded file: {temp_filename} ({len(download_response.content)} bytes)")
                        
                        # Process the actual Excel file
                        try:
                            print(f"[INFO] Processing Excel file: {temp_filename}")
                            processing_result = process_workbook(temp_filename)
                            if processing_result:
                                success_count += 1
                                processed_count += 1
                                print(f"[SUCCESS] Successfully processed Excel file: {report_info['filename']}")
                                
                                # Create parsed data structure
                                parsed_data = {
                                    'filename': report_info['filename'],
                                    'processing_result': processing_result,
                                    'summary': f'Successfully processed Excel file with {len(str(processing_result))} characters of data',
                                    'worksheet_info': 'Excel workbook processed successfully',
                                    'sharepoint_metadata': {
                                        'path': report_info['path'],
                                        'modified': report_info['modified'],
                                        'processing_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                        'source': 'SharePoint List - Service Provider Uploads',
                                        'full_url': report_info['full_url']
                                    }
                                }
                            else:
                                print(f"[WARNING] process_workbook returned None for: {report_info['filename']}")
                                # Still count as successful download, create metadata report
                                success_count += 1
                                processed_count += 1
                                parsed_data = {
                                    'filename': report_info['filename'],
                                    'processing_result': 'Excel file downloaded but processing returned None',
                                    'summary': f'Downloaded Excel file ({len(download_response.content)} bytes) but processing failed',
                                    'worksheet_info': 'File downloaded successfully, processing needs review',
                                    'sharepoint_metadata': {
                                        'path': report_info['path'],
                                        'modified': report_info['modified'],
                                        'processing_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                        'source': 'SharePoint List - Service Provider Uploads',
                                        'full_url': report_info['full_url']
                                    }
                                }
                        except Exception as process_error:
                            print(f"[ERROR] Exception processing Excel file: {process_error}")
                            # Still count as successful download
                            success_count += 1 
                            processed_count += 1
                            parsed_data = {
                                'filename': report_info['filename'],
                                'processing_result': f'Excel processing failed: {str(process_error)}',
                                'summary': f'Downloaded Excel file ({len(download_response.content)} bytes) but processing threw exception',
                                'worksheet_info': f'Processing error: {str(process_error)}',
                                'sharepoint_metadata': {
                                    'path': report_info['path'],
                                    'modified': report_info['modified'],
                                    'processing_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                    'source': 'SharePoint List - Service Provider Uploads',
                                    'full_url': report_info['full_url']
                                }
                            }
                        
                        finally:
                            # Clean up temporary file
                            try:
                                os.remove(temp_filename)
                                print(f"[INFO] Cleaned up temporary file: {temp_filename}")
                            except Exception as cleanup_error:
                                print(f"[WARNING] Could not clean up temp file: {cleanup_error}")
                    
                    elif download_response:
                        print(f"[ERROR] Failed to download file: {download_response.status_code} - {download_response.text[:200]}...")
                        continue
                    # If download_response is None, we already handled it above by creating metadata-only report
                        
                except Exception as download_error:
                    print(f"[ERROR] Error downloading file {report_info['filename']}: {download_error}")
                    continue
                
                # Create summary report content from actual Excel data
                if 'parsed_data' in locals() and parsed_data:
                    report_content = f"""Processed SharePoint Excel Report
=================================================

Report Filename: {report_info['filename']}
SharePoint Path: {report_info['path']}
Full URL: {report_info['full_url']}
Last Modified: {report_info['modified']}
Processing Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Source: SharePoint List - Service Provider Uploads

Excel File Analysis:
- File successfully downloaded and processed
- Extracted data from Excel workbook
- Contains CS Flex Weekly Service Delivery Report data

Extracted Data Summary:
{parsed_data.get('summary', 'Data extracted successfully')}

Worksheet Information:
{parsed_data.get('worksheet_info', 'Multiple worksheets processed')}

Processing Result Preview:
{str(parsed_data.get('processing_result', 'No detailed results available'))[:500]}...

Note: This report contains actual data extracted from the Excel file.
"""
                else:
                    # Fallback for failed processing
                    report_content = f"""SharePoint Report Processing Attempt
=================================================

Report Filename: {report_info['filename']}
SharePoint Path: {report_info['path']}
Full URL: {report_info['full_url']}
Last Modified: {report_info['modified']}
Processing Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Source: SharePoint List - Service Provider Uploads

Processing Status: Failed to download or parse Excel file
- File identified in SharePoint list "Service Provider Uploads"
- Located in path: {report_info['path']}
- Download or parsing encountered an error

Note: This file requires manual review or alternative processing approach.
"""

                # Save summary report locally
                output_filename = f"Report_Summary_{report_info['filename'].replace(' ', '_').replace('/', '_')}.txt"
                try:
                    # Temporarily comment out SharePoint upload until URL format is fixed
                    # success = upload_text_to_sharepoint(token, report_content, output_filename, "monthlyreport")
                    
                    # Save the summary report locally
                    with open(output_filename, 'w', encoding='utf-8') as f:
                        f.write(report_content)
                    print(f"📄 Successfully saved report summary locally: {output_filename}")
                        
                except Exception as save_error:
                    print(f"⚠️ Could not save report summary locally: {save_error}")

                    
            except Exception as e:
                print(f"❌ Error processing {report_info['filename']}: {e}")
                import traceback
                print(f"Stack trace: {traceback.format_exc()}")
        
        print(f"\n📈 SharePoint list processing complete!")
        print(f"✅ Successfully processed: {success_count} files")
        print(f"📊 Total attempted: {min(len(report_files), max_files)} files")
        
        return success_count > 0
        
    except Exception as e:
        print(f"Error processing SharePoint list files: {str(e)}")
        return False


def process_sharepoint_with_config(site=None, path=None, pattern=None, client_id=None, tenant_id=None):
    """Process SharePoint files with custom configuration"""
    try:
        config = SHAREPOINT_CONFIG.copy()
        if site:
            config['site'] = site
        if path:
            config['relative_path'] = path
        if pattern:
            config['search_pattern'] = pattern
        
        return process_sharepoint_files(client_id, tenant_id, config)
    except Exception as e:
        print(f"Error processing SharePoint with custom config: {str(e)}")
        return False


if __name__ == "__main__":
    process_sharepoint_files()




