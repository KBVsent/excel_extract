"""
Complete solution to convert all Excel sheets to images
Workaround for Spire.XLS limitation (only first 3 sheets)

Process:
1. Split Excel file by sheets (each sheet becomes a separate file)
2. Convert each split file to image
3. Clean up temporary files

This allows converting unlimited sheets to images
"""

from spire.xls import *
from spire.xls.common import *
import os
import shutil
import tempfile


def split_excel_by_sheets(input_file, output_dir):
    """
    Split Excel file by sheets
    
    Args:
        input_file: Input Excel file path
        output_dir: Output directory for split files
        
    Returns:
        List of (sheet_name, file_path) tuples
    """
    os.makedirs(output_dir, exist_ok=True)
    
    workbook = Workbook()
    workbook.LoadFromFile(input_file)
    
    split_files = []
    
    for worksheet in workbook.Worksheets:
        sheet_name = worksheet.Name
        
        # Create new workbook for this sheet
        newWorkbook = Workbook()
        newWorkbook.Worksheets.Clear()
        newWorkbook.Worksheets.AddCopy(worksheet)
        
        # Save to file
        output_file = os.path.join(output_dir, f"{sheet_name}.xlsx")
        newWorkbook.SaveToFile(output_file, FileFormat.Version2016)
        
        split_files.append((sheet_name, output_file))
        newWorkbook.Dispose()
    
    workbook.Dispose()
    
    return split_files


def convert_worksheet_to_image_no_margin(excel_file, output_file, sheet_index=0):
    """
    Convert Excel worksheet to image without margins
    
    Args:
        excel_file: Input Excel file path
        output_file: Output image file path
        sheet_index: Worksheet index (default: 0)
    """
    workbook = Workbook()
    workbook.LoadFromFile(excel_file)
    
    sheet = workbook.Worksheets.get_Item(sheet_index)
    
    # Set all margins to zero
    sheet.PageSetup.TopMargin = 0
    sheet.PageSetup.BottomMargin = 0
    sheet.PageSetup.LeftMargin = 0
    sheet.PageSetup.RightMargin = 0
    
    # Convert to image
    image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
    
    # Save image
    image.Save(output_file)
    
    workbook.Dispose()


def convert_excel_to_images(input_file, output_dir, no_margin=True, keep_temp_files=False):
    """
    Convert all sheets in Excel file to images
    Workaround for Spire.XLS limitation (only first 3 sheets)
    
    Args:
        input_file: Input Excel file path
        output_dir: Output directory for images
        no_margin: Remove margins from images (default: True)
        keep_temp_files: Keep temporary split Excel files (default: False)
        
    Returns:
        List of (sheet_name, image_path) tuples
    """
    try:
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
        # Create temporary directory for split files
        temp_dir = tempfile.mkdtemp(prefix="excel_split_")
        
        print("=" * 70)
        print(f"Converting Excel to Images: {input_file}")
        print("=" * 70)
        print()
        
        # Step 1: Split Excel by sheets
        print("Step 1: Splitting Excel file by sheets...")
        print("-" * 70)
        split_files = split_excel_by_sheets(input_file, temp_dir)
        print(f"✓ Split into {len(split_files)} sheet(s)\n")
        
        # Step 2: Convert each split file to image
        print("Step 2: Converting each sheet to image...")
        print("-" * 70)
        
        image_files = []
        
        for i, (sheet_name, excel_path) in enumerate(split_files, 1):
            # Create safe filename for image
            safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in sheet_name)
            image_path = os.path.join(output_dir, f"{safe_name}.png")
            
            try:
                # Convert to image
                if no_margin:
                    convert_worksheet_to_image_no_margin(excel_path, image_path)
                else:
                    # Use original convert function
                    workbook = Workbook()
                    workbook.LoadFromFile(excel_path)
                    sheet = workbook.Worksheets.get_Item(0)
                    image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, 
                                        sheet.LastRow, sheet.LastColumn)
                    image.Save(image_path)
                    workbook.Dispose()
                
                image_files.append((sheet_name, image_path))
                print(f"  [{i}/{len(split_files)}] ✓ '{sheet_name}' -> {image_path}")
                
            except Exception as e:
                print(f"  [{i}/{len(split_files)}] ✗ '{sheet_name}' - Error: {str(e)}")
        
        print()
        
        # Step 3: Clean up temporary files
        if not keep_temp_files:
            print("Step 3: Cleaning up temporary files...")
            print("-" * 70)
            shutil.rmtree(temp_dir)
            print(f"✓ Removed temporary directory: {temp_dir}\n")
        else:
            print(f"Step 3: Temporary files kept at: {temp_dir}\n")
        
        # Summary
        print("=" * 70)
        print(f"✓ Conversion complete!")
        print(f"  Total sheets: {len(split_files)}")
        print(f"  Successfully converted: {len(image_files)}")
        print(f"  Failed: {len(split_files) - len(image_files)}")
        print(f"  Output directory: {output_dir}")
        print("=" * 70)
        
        return image_files
        
    except Exception as e:
        print(f"\n✗ Conversion failed: {str(e)}")
        return []


def convert_excel_to_images_simple(input_file, output_dir="output/images"):
    """
    Simple wrapper function to convert Excel to images
    
    Args:
        input_file: Input Excel file path
        output_dir: Output directory for images (default: "output/images")
        
    Returns:
        List of generated image paths
    """
    result = convert_excel_to_images(input_file, output_dir, no_margin=True)
    return [img_path for _, img_path in result]


if __name__ == "__main__":
    import sys
    
    # Check command line arguments
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        output_dir = sys.argv[2] if len(sys.argv) > 2 else "output/images"
    else:
        # Default example
        input_file = "examples/示例.xlsx"
        output_dir = "output/images"
    
    # Convert Excel to images
    print()
    result = convert_excel_to_images(
        input_file=input_file,
        output_dir=output_dir,
        no_margin=True,          # Remove margins
        keep_temp_files=False    # Clean up temporary files
    )
    
    print()
    
    # Display results
    if result:
        print("Generated images:")
        for sheet_name, image_path in result:
            print(f"  • {sheet_name}: {image_path}")
    else:
        print("No images were generated.")
    
    print()
    print("Usage:")
    print(f"  python {os.path.basename(__file__)} <input_file> [output_dir]")
    print()
    print("Example:")
    print(f"  python {os.path.basename(__file__)} my_excel.xlsx output/my_images")
    print()
