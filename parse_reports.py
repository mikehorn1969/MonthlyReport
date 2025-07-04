import sys
import os
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

        
        print(f"\nProcessing file: {filename}")
        # Load the workbook
        wb = load_workbook(filename)
        
        # Get the active sheet
        sheet = wb.active
        
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

if __name__ == "__main__":
    arguments = sys.argv
    if len(arguments) < 2:
        print("Error: Usage: python parse_reports.py [dirname], where dirname contains the weekly reports") 
        sys.exit(1)

    print(f"Processing: {arguments[1]}")
    process_directory(arguments[1])




