import os
import pandas as pd

def list_all_files(folder):
    """List all files in the local folder and its subfolders."""
    files = []
    if os.path.exists(folder):
        for root, _, filenames in os.walk(folder):
            print(f"Scanning directory: {root}")
            for filename in filenames:
                full_path = os.path.join(root, filename)
                files.append(full_path)
                print(f"Found file: {full_path}")
    else:
        print(f"The folder {folder} does not exist.")
    print(f"Found {len(files)} total files.")
    return files

def list_excel_files(files):
    """Filter and list all Excel files from the list of files."""
    excel_files = [f for f in files if f.endswith('.xlsx')]
    print(f"Found {len(excel_files)} Excel files.")
    return excel_files

def read_excel_column(file_path, column_index):
    """Read a specific column from the first sheet of an Excel file."""
    try:
        df = pd.read_excel(file_path, sheet_name=0, header=None)  # Read without headers
        if column_index < df.shape[1]:  # Check if the column index is valid
            return df[[column_index]]
        else:
            print(f"Column index {column_index} not found in {file_path}")
            return pd.DataFrame()
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return pd.DataFrame()

def main():
    folder = r'C:\Users\Zakhar\Desktop\PROPOSALS'  # Specify the root folder on your PC
    column_index = 1  # Specify the column index to read (0-based index, so 'B' is 1)
    print(f"Folder to scan: {folder}")
    
    all_files = list_all_files(folder)
    
    if all_files:
        print("All files found:")
        for file in all_files:
            print(file)
    
    excel_files = list_excel_files(all_files)
    
    if excel_files:
        print("Excel files found:")
        for file in excel_files:
            print(file)
    
        all_data = pd.DataFrame()
    
        for file in excel_files:
            df = read_excel_column(file, column_index)
            if not df.empty:
                print(f"Reading file: {file}")
                print(df.head())  # Print first few rows of the DataFrame
                all_data = pd.concat([all_data, df], ignore_index=True)

        # Remove duplicates
        all_data.drop_duplicates(inplace=True)
        
        # Check the data before saving
        print("Combined Data:")
        print(all_data.head())

        # Save the combined data to a new Excel file
        if not all_data.empty:
            all_data.to_excel('combined_unique_column.xlsx', index=False)
            print("Data extraction and duplication removal complete. File saved as 'combined_unique_column.xlsx'.")
        else:
            print("No data to save.")
    else:
        print("No Excel files found.")

if __name__ == "__main__":
    main()
