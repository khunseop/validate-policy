import xlwings as xw
import pandas as pd

def parse_policy_file(file_path):
    """
    Parses a firewall policy Excel file, extracts 'Rulename' and 'Enable' columns,
    removes duplicates, and strips whitespace from the data.
    
    Handles files with many blank cells where 'Rulename' and 'Enable' exist 
    at the top of each row with spaces before them.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame: A DataFrame with 'Rulename' and 'Enable' columns,
                      with duplicates removed and whitespace stripped.
    """
    try:
        # Use xlwings to open the workbook and read the data
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            # Assuming data is in the first sheet
            ws = wb.sheets[0]

            # Read all data from the sheet (without assuming header position)
            # First, find the header row by searching for 'Rulename' and 'Enable'
            max_row = ws.used_range.last_cell.row if ws.used_range else 1000
            max_col = ws.used_range.last_cell.column if ws.used_range else 100
            
            # Find header row by searching for 'Rulename' and 'Enable' columns
            header_row = None
            rulename_col_idx = None
            enable_col_idx = None
            
            # Search for header row (check first 20 rows)
            for row_idx in range(1, min(21, max_row + 1)):
                for col_idx in range(1, min(max_col + 1, 100)):
                    cell_value = ws.range((row_idx, col_idx)).value
                    if cell_value:
                        cell_str = str(cell_value).strip()
                        if cell_str.lower() == 'rulename' and rulename_col_idx is None:
                            rulename_col_idx = col_idx
                        if cell_str.lower() == 'enable' and enable_col_idx is None:
                            enable_col_idx = col_idx
                
                # If we found both columns, this is the header row
                if rulename_col_idx and enable_col_idx:
                    header_row = row_idx
                    break
            
            if header_row is None or rulename_col_idx is None or enable_col_idx is None:
                raise ValueError(f"Could not find 'Rulename' and 'Enable' columns in {file_path}")
            
            # Now extract data rows (starting from header_row + 1)
            data_rows = []
            for row_idx in range(header_row + 1, max_row + 1):
                rulename_value = ws.range((row_idx, rulename_col_idx)).value
                enable_value = ws.range((row_idx, enable_col_idx)).value
                
                # Skip rows where both values are empty/None
                if rulename_value is None and enable_value is None:
                    continue
                
                # Convert to string and strip whitespace
                rulename_str = str(rulename_value).strip() if rulename_value is not None else ''
                enable_str = str(enable_value).strip() if enable_value is not None else ''
                
                # Skip rows where both are empty strings after stripping
                if not rulename_str and not enable_str:
                    continue
                
                data_rows.append({
                    'Rulename': rulename_str,
                    'Enable': enable_str
                })
            
            wb.close()

        # Create DataFrame from extracted data
        if not data_rows:
            return pd.DataFrame(columns=['Rulename', 'Enable'])
        
        df = pd.DataFrame(data_rows)
        
        # Remove duplicate rows
        df_processed = df.drop_duplicates()
        
        # Remove rows where both columns are empty
        df_processed = df_processed[
            ~((df_processed['Rulename'] == '') & (df_processed['Enable'] == ''))
        ]

        return df_processed

    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame() # Return empty DataFrame on error

if __name__ == "__main__":
    running_policy_file = "running_policy.xlsx"
    candidate_policy_file = "candidate_policy.xlsx"
    
    running_output_file = "running_policy_processed.xlsx"
    candidate_output_file = "candidate_policy_processed.xlsx"
    
    print(f"--- Parsing Running Policy: {running_policy_file} ---")
    running_policy_data = parse_policy_file(running_policy_file)
    print(running_policy_data)
    if not running_policy_data.empty:
        running_policy_data.to_excel(running_output_file, index=False)
        print(f"Processed running policy saved to {running_output_file}")
    print("\n" + "="*50 + "\n")

    print(f"--- Parsing Candidate Policy: {candidate_policy_file} ---")
    candidate_policy_data = parse_policy_file(candidate_policy_file)
    print(candidate_policy_data)
    if not candidate_policy_data.empty:
        candidate_policy_data.to_excel(candidate_output_file, index=False)
        print(f"Processed candidate policy saved to {candidate_output_file}")
    print("\n" + "="*50 + "\n")
