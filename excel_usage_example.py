"""
Example of how to use parse_reports.py from Excel Python environment

This file demonstrates how to call the report parsing functions
from within Excel using Python in Excel.
"""

# Example usage in Excel Python environment:

# Method 1: Process current folder
# import parse_reports
# result = parse_reports.process_current_folder()
# print(f"Processing result: {result}")

# Method 2: Process specific folder
# import parse_reports
# folder_path = "C:\\Users\\MikeHorn\\workspace\\MonthlyReport"
# result = parse_reports.process_folder(folder_path)
# print(f"Processing result: {result}")

# Method 3: Use main function (automatically detects environment)
# import parse_reports
# result = parse_reports.main()
# print(f"Processing result: {result}")

if __name__ == "__main__":
    # This will run when executed directly (not from Excel)
    import parse_reports
    print("Testing parse_reports from Python script...")
    result = parse_reports.process_current_folder()
    print(f"Test completed. Success: {result}")