"""
SharePoint Configuration for parse_reports.py

This file contains the SharePoint configuration settings.
Modify these values to match your SharePoint environment.
"""

from datetime import datetime, timedelta
import calendar
from keyvault import get_secret

# SharePoint Site Configuration
SHAREPOINT_SITE = "jjag.sharepoint.com"

# Dynamic SharePoint Document Library Path (automatically uses previous month)
def get_previous_month_path():
    """Generate SharePoint path for the previous month"""
    # Get current date
    today = datetime.now()
    
    # Calculate first day of current month, then subtract 1 day to get last day of previous month
    first_day_current_month = today.replace(day=1)
    last_day_previous_month = first_day_current_month - timedelta(days=1)
    
    # Get previous month details
    prev_year = last_day_previous_month.year
    prev_month_num = last_day_previous_month.month
    prev_month_name = calendar.month_name[prev_month_num]
    
    # Format: "/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/YYYY/MM - MonthName"
    return f"/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/{prev_year}/{prev_month_num:02d} - {prev_month_name}"

def get_current_month_path():
    """Generate SharePoint path for the current month"""
    # Get current date
    today = datetime.now()
    
    # Get current month details
    curr_year = today.year
    curr_month_num = today.month
    curr_month_name = calendar.month_name[curr_month_num]
    
    # Format: "/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/YYYY/MM - MonthName"
    return f"/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/{curr_year}/{curr_month_num:02d} - {curr_month_name}"
    
def get_specific_month_path(year, month):
    """Generate SharePoint path for a specific year and month"""
    month_name = calendar.month_name[month]
    return f"/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/{year}/{month:02d} - {month_name}"

# Get dynamic path (defaults to previous month)
# over-engineered to allow easy switching between current, previous, or specific months instead of searching root, which is what we do now
SHAREPOINT_PATH = get_previous_month_path()
#SHAREPOINT_PATH = "/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients"

# File Search Pattern
SEARCH_PATTERN = "CS Flex Weekly Service Delivery Report"

# Tech Debt: these should be stored in a secure vault or environment variables MONTHLYREPORT-CLIENTID, MONTHLYREPORT-TENANTID
# Azure AD Configuration 
try:
    CLIENT_ID = get_secret("MONTHLYREPORT-CLIENTID", default_value="4358f4b7-e580-4ebb-b105-f561110e0b85")
    TENANT_ID = get_secret("MONTHLYREPORT-TENANTID", default_value="44ff6ede-31e4-4f2b-b73b-b30140741c4f")
except Exception:
    # Fallback if keyvault is not available
    CLIENT_ID = "4358f4b7-e580-4ebb-b105-f561110e0b85"
    TENANT_ID = "44ff6ede-31e4-4f2b-b73b-b30140741c4f"
# Redirect URI for MSAL authentication
# This MUST be registered in your Azure AD App Registration under Authentication
REDIRECT_URI = "http://localhost:8080"

# Alternative paths for different clients or time periods
ALTERNATIVE_CONFIGS = {
    "current_month": {
        "site": "jjag.sharepoint.com",
        "path": get_current_month_path(),
        "pattern": "CS Flex Weekly Service Delivery Report"
    },
    "previous_month": {
        "site": "jjag.sharepoint.com",
        "path": get_previous_month_path(),
        "pattern": "CS Flex Weekly Service Delivery Report"
    },
    "root": {
        "site": "jjag.sharepoint.com",
        "path": "/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/monthlyreport",
        "pattern": "CS Flex Weekly Service Delivery Report"
    }
    # Add more configurations as needed
}

def get_config(config_name=None):
    """Get SharePoint configuration by name or default"""
    if config_name and config_name in ALTERNATIVE_CONFIGS:
        return ALTERNATIVE_CONFIGS[config_name]
    else:
        return {
            "site": SHAREPOINT_SITE,
            "relative_path": SHAREPOINT_PATH,
            "search_pattern": SEARCH_PATTERN
        }