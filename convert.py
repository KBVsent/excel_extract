import pandas as pd
import re
from pathlib import Path

# ============================================================================
# CONFIGURATION - Edit these parameters
# ============================================================================

# Input Excel file path
INPUT_FILE = "examples/genexus.xlsx"  # Change this to your Excel file path

# Output Markdown file path (None for auto-generated name, ignored if ENABLE_PAGINATION=True)
OUTPUT_FILE = None

# Clean mode: 'auto', 'aggressive', 'minimal', 'none'
CLEAN_MODE = 'auto'

# Pagination options
ENABLE_PAGINATION = True  # Set to True to save each sheet as a separate .md file
OUTPUT_FOLDER = "output"   # Folder for paginated output (used when ENABLE_PAGINATION=True)

# Additional options
SKIP_EMPTY_ROWS = True
SKIP_UNNAMED_COLS = True

# ============================================================================
# END CONFIGURATION
# ============================================================================


def clean_dataframe(df, mode='auto'):
    """
    Clean dataframe by removing NaN and unnamed columns
    
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


def convert_excel_to_markdown(
    input_file: str,
    output_file: str = None,
    clean_mode: str = 'auto',
    enable_pagination: bool = False,
    output_folder: str = "output"
):
    """
    Convert Excel file to Markdown format
    
    Args:
        input_file: Path to input Excel file
        output_file: Path to output Markdown file (optional, ignored if enable_pagination=True)
        clean_mode: Cleaning mode
        enable_pagination: Save each sheet to a separate .md file
        output_folder: Folder for paginated output files
    """
    # Validate input file
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    if input_path.suffix.lower() not in ['.xlsx', '.xls']:
        raise ValueError(f"Unsupported file format: {input_path.suffix}")
    
    print(f"Converting: {input_path}")
    print(f"Clean mode: {clean_mode}")
    
    # Set output path(s)
    if enable_pagination:
        output_dir = Path(output_folder)
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Pagination: Enabled")
        print(f"Output folder: {output_dir}")
    else:
        if output_file is None:
            output_path = input_path.with_suffix('.md')
        else:
            output_path = Path(output_file)
        print(f"Pagination: Disabled")
        print(f"Output: {output_path}")
    
    print()
    
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(input_path)
        
        if enable_pagination:
            # Pagination mode: save each sheet to separate file
            output_files = []
            
            for sheet_name in excel_file.sheet_names:
                print(f"Processing sheet: {sheet_name}")
                
                # Read sheet without header
                df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
                
                # Clean dataframe
                df = clean_dataframe(df, mode=clean_mode)
                
                # Skip if empty after cleaning
                if df.empty:
                    print(f"  → Skipped (empty after cleaning)")
                    continue
                
                # Create safe filename from sheet name
                safe_filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in sheet_name)
                safe_filename = safe_filename.strip().replace(' ', '_')
                sheet_output_path = output_dir / f"{safe_filename}.md"
                
                # Build markdown content
                markdown_content = []
                markdown_content.append(f"# {sheet_name}\n")
                
                # Use first row as header if it looks like a header
                first_row = df.iloc[0]
                if all(isinstance(val, str) or val != '' for val in first_row):
                    df.columns = first_row
                    df = df.iloc[1:]
                
                # Convert to markdown table
                markdown_table = df.to_markdown(index=False)
                markdown_content.append(markdown_table)
                
                # Write to individual file
                final_content = '\n'.join(markdown_content)
                sheet_output_path.write_text(final_content, encoding='utf-8')
                output_files.append(sheet_output_path)
                
                print(f"  → Saved to: {sheet_output_path.name}")
            
            print()
            print("✅ Conversion completed!")
            print(f"Created {len(output_files)} files in: {output_dir}")
            
        else:
            # Single file mode: combine all sheets
            markdown_content = []
            
            for sheet_name in excel_file.sheet_names:
                print(f"Processing sheet: {sheet_name}")
                
                # Read sheet without header
                df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
                
                # Clean dataframe
                df = clean_dataframe(df, mode=clean_mode)
                
                # Skip if empty after cleaning
                if df.empty:
                    print(f"  → Skipped (empty after cleaning)")
                    continue
                
                # Add sheet name as header
                markdown_content.append(f"## {sheet_name}\n")
                
                # Use first row as header if it looks like a header
                first_row = df.iloc[0]
                if all(isinstance(val, str) or val != '' for val in first_row):
                    df.columns = first_row
                    df = df.iloc[1:]
                
                # Convert to markdown table
                markdown_table = df.to_markdown(index=False)
                markdown_content.append(markdown_table)
                markdown_content.append("\n")
                
                print(f"  → Converted successfully")
            
            # Write to single file
            final_content = '\n'.join(markdown_content)
            output_path.write_text(final_content, encoding='utf-8')
            
            print()
            print("✅ Conversion completed!")
            print(f"Output saved to: {output_path}")
        
    except Exception as e:
        print(f"❌ Conversion failed: {e}")
        raise


def main():
    """Main function"""
    print("="*70)
    print("Excel to Markdown Converter")
    print("="*70)
    print()
    
    convert_excel_to_markdown(
        input_file=INPUT_FILE,
        output_file=OUTPUT_FILE,
        clean_mode=CLEAN_MODE,
        enable_pagination=ENABLE_PAGINATION,
        output_folder=OUTPUT_FOLDER
    )


if __name__ == '__main__':
    main()
