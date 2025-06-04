import sys
from openpyxl import load_workbook

def is_cell_merged(sheet, cell_coord):
    """Check if a specific cell is part of a merged range"""
    for merged_range in sheet.merged_cells.ranges:
        if cell_coord in merged_range:
            return True, merged_range
    return False, None

def unmerge_cells_in_range(filename):
    try:
        # Load the workbook
        wb = load_workbook(filename)
        
        # Get the active sheet
        sheet = wb.active
        
        # First check if E31 is merged
        is_merged, target_range = is_cell_merged(sheet, "E31")
        
        if is_merged:
            print("Cell E31 is not merged. No action taken.")
            return
            
            # Unmerge cells in the specified range
            for merged_cell_range in list(sheet.merged_cells.ranges):
                min_col, min_row, max_col, max_row = merged_cell_range.bounds
                # Check if the merged cell range overlaps with our target range (D31:R42)
                if not (max_col < 4 or min_col > 18 or max_row < 31 or min_row > 42):
                    sheet.unmerge_cells(str(merged_cell_range))

        # Save the workbook
            wb.save(filename)
            print(f"Successfully unmerged cells in range D31:R42 in {filename}")

        else: #is_merged

            print("Cell E31 is not merged. No action taken.")

        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python parse_reports.py <excel_file>")
        sys.exit(1)
    
    filename = sys.argv[1]
    unmerge_cells_in_range(filename)




