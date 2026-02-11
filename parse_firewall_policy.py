import xlwings as xw
import pandas as pd

def parse_policy_file(file_path):
    """
    Parses a firewall policy Excel file, extracts 'Rulename' and 'Enable' columns,
    removes duplicates, and strips whitespace from the data.

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

            # Read the entire sheet into a pandas DataFrame
            # header=1 means the second row is the header (0-indexed)
            df = ws.range('A1').expand('table').options(pd.DataFrame, header=1).value
            wb.close()

        # Select 'Rulename' and 'Enable' columns
        # Ensure column names are stripped of any potential leading/trailing whitespace
        df.columns = df.columns.str.strip()
        required_columns = ['Rulename', 'Enable']

        # Check if all required columns exist
        if not all(col in df.columns for col in required_columns):
            missing_cols = [col for col in required_columns if col not in df.columns]
            raise ValueError(f"Missing required columns in {file_path}: {missing_cols}")

        df_filtered = df[required_columns].copy()

        # Fill NaN values with empty strings before stripping whitespace from string columns
        for col in df_filtered.columns:
            df_filtered[col] = df_filtered[col].fillna('').astype(str).str.strip()

        # Remove duplicate rows
        df_processed = df_filtered.drop_duplicates()

        return df_processed

    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
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
    print("
" + "="*50 + "
")

    print(f"--- Parsing Candidate Policy: {candidate_policy_file} ---")
    candidate_policy_data = parse_policy_file(candidate_policy_file)
    print(candidate_policy_data)
    if not candidate_policy_data.empty:
        candidate_policy_data.to_excel(candidate_output_file, index=False)
        print(f"Processed candidate policy saved to {candidate_output_file}")
    print("
" + "="*50 + "
")
