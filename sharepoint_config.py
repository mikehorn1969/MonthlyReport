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
# Get dynamic path (defaults to previous month)
# over-engineered to allow easy switching between current, previous, or specific months instead of searching root, which is what we do now

# File Search Pattern
SEARCH_PATTERN = "CS Flex Weekly Service Delivery Report"
CLIENT_ID = get_secret("MONTHLYREPORT-CLIENTID")
TENANT_ID = get_secret("MONTHLYREPORT-TENANTID")
CLIENT_SECRET = get_secret("MONTHLYREPORT-CLIENTSECRET")


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
    return f"{prev_year}/{prev_month_num:02d} - {prev_month_name}"


def get_current_month_path():
    """Generate SharePoint path for the current month"""
    # Get current date
    today = datetime.now()
    
    # Get current month details
    curr_year = today.year
    curr_month_num = today.month
    curr_month_name = calendar.month_name[curr_month_num]
    
    # Format: "/sites/InternalTeam/Shared Documents/Restricted/Clients/Julian Brown - Clients/YYYY/MM - MonthName"
    return f"{curr_year}/{curr_month_num:02d} - {curr_month_name}"


def get_specific_month_path(year, month):
    """Generate SharePoint path for a specific year and month"""
    month_name = calendar.month_name[month]
    return f"{year}/{month:02d} - {month_name}"
