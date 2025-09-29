"""
Excel Report Parser

This script parses Excel weekly reports and extracts information into text files.

Usage from Excel (Python in Excel):
1. Import this module: import parse_reports
2. Process current folder: parse_reports.process_current_folder()
3. Process specific folder: parse_reports.process_folder("C:\\path\\to\\folder")

Usage from command line:
python parse_reports.py [directory_path]

Functions available for Excel:
- process_current_folder(): Process Excel files in current working directory
- process_folder(path): Process Excel files in specified directory
- main(): Main function (handles both command line and Excel usage)
"""

import sys
import os
from pathlib import Path
from openpyxl import load_workbook

def is_cell_merged(sheet, cell_coord):
    """Check if a specific cell is part of a merged range"""
    for merged_range in sheet.merged_cells.ranges:
        if cell_coord in merged_range:
            return True, merged_range
    return False, None

def process_workbook(filename):
    """Process a single workbook and extract information"""
    try:
        # Ensure we have absolute path
        filename = os.path.abspath(filename)
        print(f"\nProcessing file: {filename}")
        
        # Check if file exists
        if not os.path.exists(filename):
            print(f"Error: File '{filename}' does not exist")
            return False
            
        # Load the workbook
        wb = load_workbook(filename, data_only=True)
        
        # Get the active sheet
        sheet = wb.active
        
        # Validate that we have a sheet
        if sheet is None:
            print(f"Error: Could not access active sheet in {filename}")
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
        
        name, ext = filename.rsplit('.', 1)  # Split on last dot
        fname = f"{name}.txt"

        with open(fname, 'w') as file:
    
            # Print contents of specific cells
            file.write(f"Week Ending: {sheet['G7'].value}\n")
            file.write(f"Service Provider: {sheet['G11'].value}\n")
            file.write(f"Client: {sheet['G13'].value}\n")

            file.write("\nService Standard updates:\n")
            file.write("SSN|Status|Comments\n")
            
            # Service standards
            for row in range(34, 43):
                if sheet[f'D{row}'].value:
                    file.write(f"{sheet[f'D{row}'].value}|{sheet[f'J{row}'].value}|{sheet[f'K{row}'].value}\n")

            # Service Risks
            file.write("\nService Risks:\n")
            file.write("Risk No|Description|Likelihood|Impact|Mitigation\n")

            for row in range(45, 48):
                if sheet[f'D{row}'].value:
                    file.write(f"{sheet[f'D{row}'].value}|{sheet[f'E{row}'].value}|{sheet[f'H{row}'].value}|{sheet[f'J{row}'].value}|{sheet[f'K{row}'].value}\n")

            #Service Issues
            file.write("\nService Issues:\n")
            file.write("Issue No|Description|Impact|Mitigation\n")

            for row in range(50, 53):
                if sheet[f'D{row}'].value:
                    file.write(f"{sheet[f'D{row}'].value}|{sheet[f'E{row}'].value}|{sheet[f'J{row}'].value}|{sheet[f'K{row}'].value}\n")

            # Planned Activities
            file.write("\nPlanned Activities:\n")
            file.write(f"{sheet['D57'].value}")


            # Client Updates
            file.write("\nClient Updates:\n")
            file.write(f"{sheet['D67'].value}")

        # Close workbook without saving changes
        wb.close()
        print("Workbook closed without saving changes")
        return True

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

def main():
    """Main function that can be called from Excel or command line"""
    try:
        # Check if running from command line with arguments
        if len(sys.argv) >= 2:
            directory_path = sys.argv[1]
            print(f"Processing directory from command line: {directory_path}")
        else:
            # Get current working directory when running from Excel
            directory_path = os.getcwd()
            print(f"Processing current directory: {directory_path}")
        
        # Process the directory
        success = process_directory(directory_path)
        
        if success:
            print("Processing completed successfully")
        else:
            print("Processing failed")
            
        return success
        
    except Exception as e:
        print(f"Error in main function: {str(e)}")
        return False

if __name__ == "__main__":
    main()




