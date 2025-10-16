"""
Excel to Markdown Converter using MarkItDown

This script uses Microsoft's MarkItDown library to convert Excel files to Markdown format.
MarkItDown provides better document structure preservation and is optimized for LLM consumption.

Repository: https://github.com/microsoft/markitdown
"""

from markitdown import MarkItDown
from pathlib import Path
import re
import pandas as pd

# ============================================================================
# CONFIGURATION - Edit these parameters
# ============================================================================

# Input Excel file path
INPUT_FILE = "examples/genexus.xlsx"  # Change this to your Excel file path

# Output Markdown file path (None for auto-generated name)
OUTPUT_FILE = None

# Pagination options
ENABLE_PAGINATION = True  # Set to True to save each sheet as a separate .md file
OUTPUT_FOLDER = "output"   # Folder for paginated output (used when ENABLE_PAGINATION=True)

# MarkItDown options
ENABLE_PLUGINS = False  # Enable 3rd-party plugins if needed

# Clean mode: 'auto', 'aggressive', 'minimal', 'none'
CLEAN_MODE = 'auto'  # Clean NaN and Unnamed columns after conversion

# ============================================================================
# END CONFIGURATION
# ============================================================================


def clean_dataframe(df, mode='auto'):
    """
    Clean dataframe by removing NaN and unnamed columns
    Same logic as convert.py for consistency
    
    Args:
        df: pandas DataFrame
        mode: cleaning mode ('auto', 'aggressive', 'minimal', 'none')
    
    Returns:
        Cleaned DataFrame
    """
    if mode == 'none':
        return df
    
    # Remove completely empty rows
    df = df.dropna(how='all')
    
    # Remove completely empty columns
    df = df.dropna(axis=1, how='all')
    
    if mode in ['auto', 'aggressive']:
        # Remove columns with Unnamed pattern
        unnamed_pattern = re.compile(r'^Unnamed:')
        cols_to_keep = []
        for col in df.columns:
            if not unnamed_pattern.match(str(col)):
                cols_to_keep.append(col)
            else:
                # Check if this unnamed column has meaningful data
                non_null_count = df[col].notna().sum()
                if non_null_count > len(df) * 0.5:  # More than 50% non-null
                    cols_to_keep.append(col)
        
        if cols_to_keep:
            df = df[cols_to_keep]
        
        # Replace NaN with empty string
        df = df.fillna('')
        
        # Remove rows where all values are empty strings
        df = df.loc[~(df == '').all(axis=1)]
    
    if mode == 'aggressive':
        # Remove sparse rows (rows with too many empty cells)
        threshold = min(3, len(df.columns) * 0.3)
        df = df.loc[df.apply(lambda x: (x != '').sum() >= threshold, axis=1)]
    
    return df


def process_sheet_with_pandas(input_file: str, sheet_name: str, clean_mode: str) -> str:
    """
    Process single Excel sheet using pandas for cleaning
    
    Args:
        input_file: Path to Excel file
        sheet_name: Name of the sheet
        clean_mode: Cleaning mode
    
    Returns:
        Cleaned markdown content
    """
    # Read sheet without header
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
    
    # Clean dataframe using the same logic as convert.py
    df = clean_dataframe(df, mode=clean_mode)
    
    # Skip if empty after cleaning
    if df.empty:
        return None
    
    # Use first row as header if it looks like a header
    first_row = df.iloc[0]
    if all(isinstance(val, str) or val != '' for val in first_row):
        df.columns = first_row
        df = df.iloc[1:]
    
    # Convert to markdown table
    markdown_table = df.to_markdown(index=False)
    
    return markdown_table


def convert_excel_to_markdown_single_file(
    input_file: str,
    output_file: str = None,
    enable_plugins: bool = False,
    clean_mode: str = 'auto'
):
    """
    Convert Excel file to a single Markdown file using MarkItDown
    
    Args:
        input_file: Path to input Excel file
        output_file: Path to output Markdown file (optional)
        enable_plugins: Enable MarkItDown plugins
        clean_mode: Cleaning mode for output
    """
    input_path = Path(input_file)
    
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    if input_path.suffix.lower() not in ['.xlsx', '.xls']:
        raise ValueError(f"Unsupported file format: {input_path.suffix}")
    
    # Set output path
    if output_file is None:
        output_path = input_path.with_suffix('.md')
    else:
        output_path = Path(output_file)
    
    print(f"Converting: {input_path}")
    print(f"Clean mode: {clean_mode}")
    print(f"Output: {output_path}")
    print()
    
    try:
        # Read all sheets for processing
        excel_file = pd.ExcelFile(input_path)
        markdown_content = []
        
        print("Processing with pandas and clean logic...")
        
        for sheet_name in excel_file.sheet_names:
            print(f"  Processing sheet: {sheet_name}")
            
            # Process sheet with pandas
            sheet_markdown = process_sheet_with_pandas(str(input_path), sheet_name, clean_mode)
            
            if sheet_markdown:
                # Add sheet name as header
                markdown_content.append(f"## {sheet_name}\n")
                markdown_content.append(sheet_markdown)
                markdown_content.append("\n")
        
        # Write to file
        final_content = '\n'.join(markdown_content)
        output_path.write_text(final_content, encoding='utf-8')
        
        print()
        print("✅ Conversion completed!")
        print(f"Output saved to: {output_path}")
        
    except Exception as e:
        print(f"❌ Conversion failed: {e}")
        raise


def convert_excel_to_markdown_paginated(
    input_file: str,
    output_folder: str = "output",
    enable_plugins: bool = False,
    clean_mode: str = 'auto'
):
    """
    Convert Excel file to separate Markdown files (one per sheet) using MarkItDown
    
    Note: MarkItDown converts the entire Excel file to a single markdown output.
    This function will parse the output and split it by sheet headings.
    
    Args:
        input_file: Path to input Excel file
        output_folder: Folder for output files
        enable_plugins: Enable MarkItDown plugins
        clean_mode: Cleaning mode for output
    """
    input_path = Path(input_file)
    
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    if input_path.suffix.lower() not in ['.xlsx', '.xls']:
        raise ValueError(f"Unsupported file format: {input_path.suffix}")
    
    # Create output folder
    output_dir = Path(output_folder)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Converting: {input_path}")
    print(f"Clean mode: {clean_mode}")
    print(f"Pagination: Enabled")
    print(f"Output folder: {output_dir}")
    print()
    
    try:
        # Read all sheets for processing
        excel_file = pd.ExcelFile(input_path)
        output_files = []
        
        print("Processing with pandas and clean logic...")
        
        for sheet_name in excel_file.sheet_names:
            print(f"  Processing sheet: {sheet_name}")
            
            # Process sheet with pandas
            sheet_markdown = process_sheet_with_pandas(str(input_path), sheet_name, clean_mode)
            
            if sheet_markdown is None:
                print(f"    → Skipped (empty after cleaning)")
                continue
            
            # Create safe filename from sheet name
            safe_filename = "".join(
                c if c.isalnum() or c in (' ', '-', '_') else '_' 
                for c in sheet_name
            )
            safe_filename = safe_filename.strip().replace(' ', '_')
            
            # Build markdown content with sheet title
            markdown_content = []
            markdown_content.append(f"# {sheet_name}\n")
            markdown_content.append(sheet_markdown)
            
            # Save to file
            sheet_output_path = output_dir / f"{safe_filename}.md"
            final_content = '\n'.join(markdown_content)
            sheet_output_path.write_text(final_content, encoding='utf-8')
            output_files.append(sheet_output_path)
            
            print(f"    → Saved to: {sheet_output_path.name}")
        
        print()
        print("✅ Conversion completed!")
        print(f"Created {len(output_files)} files in: {output_dir}")
        
    except Exception as e:
        print(f"❌ Conversion failed: {e}")
        raise


def convert_excel_to_markdown(
    input_file: str,
    output_file: str = None,
    enable_pagination: bool = False,
    output_folder: str = "output",
    enable_plugins: bool = False,
    clean_mode: str = 'auto'
):
    """
    Convert Excel file to Markdown using MarkItDown
    
    Args:
        input_file: Path to input Excel file
        output_file: Path to output Markdown file (optional, ignored if enable_pagination=True)
        enable_pagination: Save each sheet to a separate .md file
        output_folder: Folder for paginated output files
        enable_plugins: Enable MarkItDown plugins
        clean_mode: Cleaning mode ('auto', 'aggressive', 'minimal', 'none')
    """
    if enable_pagination:
        convert_excel_to_markdown_paginated(
            input_file=input_file,
            output_folder=output_folder,
            enable_plugins=enable_plugins,
            clean_mode=clean_mode
        )
    else:
        convert_excel_to_markdown_single_file(
            input_file=input_file,
            output_file=output_file,
            enable_plugins=enable_plugins,
            clean_mode=clean_mode
        )


def main():
    """Main function"""
    print("=" * 70)
    print("Excel to Markdown Converter (MarkItDown)")
    print("=" * 70)
    print()
    
    convert_excel_to_markdown(
        input_file=INPUT_FILE,
        output_file=OUTPUT_FILE,
        enable_pagination=ENABLE_PAGINATION,
        output_folder=OUTPUT_FOLDER,
        enable_plugins=ENABLE_PLUGINS,
        clean_mode=CLEAN_MODE
    )


if __name__ == '__main__':
    main()
