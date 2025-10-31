"""
Use Spire.XLS to split Excel file by sheets
Split a multi-sheet Excel file into separate Excel files, one file per sheet
"""

from spire.xls import *
from spire.xls.common import *
import os


def split_excel_by_sheets(input_file, output_dir):
    """
    Split Excel file by sheets, each sheet will be saved as a separate Excel file
    
    Args:
        input_file: Input Excel file path
        output_dir: Output directory for split files
    """
    try:
        # Create output directory if not exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Create Workbook object and load Excel file
        workbook = Workbook()
        workbook.LoadFromFile(input_file)
        
        print(f"Loading Excel file: {input_file}")
        print(f"Total sheets: {workbook.Worksheets.Count}\n")
        
        # Iterate through all worksheets
        for i, worksheet in enumerate(workbook.Worksheets, 1):
            # Create a new Workbook object
            newWorkbook = Workbook()
            # Clear default worksheets in new workbook
            newWorkbook.Worksheets.Clear()
            
            # Copy worksheet from original Excel file to new workbook
            newWorkbook.Worksheets.AddCopy(worksheet)
            
            # Save new workbook to specified folder
            output_file = os.path.join(output_dir, f"{worksheet.Name}.xlsx")
            newWorkbook.SaveToFile(output_file, FileFormat.Version2016)
            
            print(f"[{i}/{workbook.Worksheets.Count}] Successfully split sheet '{worksheet.Name}' to: {output_file}")
            
            # Release resources
            newWorkbook.Dispose()
        
        # Release original workbook resources
        workbook.Dispose()
        
        print(f"\n✓ Split complete! All sheets have been saved to: {output_dir}")
        
    except Exception as e:
        print(f"✗ Split failed: {str(e)}")


def split_excel_by_sheets_with_filter(input_file, output_dir, sheet_names=None, exclude_sheets=None):
    """
    Split Excel file by sheets with filtering options
    
    Args:
        input_file: Input Excel file path
        output_dir: Output directory for split files
        sheet_names: List of sheet names to split (if None, split all sheets)
        exclude_sheets: List of sheet names to exclude from splitting
    """
    try:
        # Create output directory if not exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Create Workbook object and load Excel file
        workbook = Workbook()
        workbook.LoadFromFile(input_file)
        
        print(f"Loading Excel file: {input_file}")
        print(f"Total sheets: {workbook.Worksheets.Count}\n")
        
        split_count = 0
        
        # Iterate through all worksheets
        for worksheet in workbook.Worksheets:
            sheet_name = worksheet.Name
            
            # Skip if sheet_names is specified and current sheet is not in the list
            if sheet_names is not None and sheet_name not in sheet_names:
                print(f"Skipping sheet '{sheet_name}' (not in include list)")
                continue
            
            # Skip if current sheet is in exclude list
            if exclude_sheets is not None and sheet_name in exclude_sheets:
                print(f"Skipping sheet '{sheet_name}' (in exclude list)")
                continue
            
            # Create a new Workbook object
            newWorkbook = Workbook()
            # Clear default worksheets in new workbook
            newWorkbook.Worksheets.Clear()
            
            # Copy worksheet from original Excel file to new workbook
            newWorkbook.Worksheets.AddCopy(worksheet)
            
            # Save new workbook to specified folder
            output_file = os.path.join(output_dir, f"{sheet_name}.xlsx")
            newWorkbook.SaveToFile(output_file, FileFormat.Version2016)
            
            split_count += 1
            print(f"✓ Successfully split sheet '{sheet_name}' to: {output_file}")
            
            # Release resources
            newWorkbook.Dispose()
        
        # Release original workbook resources
        workbook.Dispose()
        
        print(f"\n✓ Split complete! {split_count} sheet(s) have been saved to: {output_dir}")
        
    except Exception as e:
        print(f"✗ Split failed: {str(e)}")


if __name__ == "__main__":
    # Example usage
    
    print("=" * 60)
    print("Excel Sheet Splitter - Using Spire.XLS")
    print("=" * 60)
    print()
    
    # Input Excel file
    input_file = "示例.xlsx"
    
    # Example 1: Split all sheets
    print("Example 1: Split all sheets")
    print("-" * 60)
    output_dir1 = "output/split_all"
    split_excel_by_sheets(input_file, output_dir1)
    print()
    
    # Example 2: Split specific sheets only
    print("\nExample 2: Split specific sheets only")
    print("-" * 60)
    output_dir2 = "output/split_filtered"
    # Only split sheets named 'Sheet1' and 'Sheet2'
    split_excel_by_sheets_with_filter(
        input_file, 
        output_dir2, 
        sheet_names=['Sheet1', 'Sheet2']
    )
    print()
    
    # Example 3: Split all sheets except certain ones
    print("\nExample 3: Split all sheets except certain ones")
    print("-" * 60)
    output_dir3 = "output/split_exclude"
    # Split all sheets except 'Sheet3'
    split_excel_by_sheets_with_filter(
        input_file, 
        output_dir3, 
        exclude_sheets=['Sheet3']
    )
    print()
    
    print("=" * 60)
    print("All operations completed!")
    print("=" * 60)
