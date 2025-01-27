import sys
import os
import pandas as pd

def process_excel(excel1_path, excel2_path, output_dir="processed"):
    try:
        # Ensure the output directory exists
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Output directory created: {output_dir}")

        # Debugging: Print the paths of the files
        print(f"Processing Excel files...\nExcel 1 Path: {excel1_path}\nExcel 2 Path: {excel2_path}")

        # Read Excel 1 and Excel 2
        df1 = pd.read_excel(excel1_path)
        df2 = pd.read_excel(excel2_path)

        # Debugging: Print the first few rows and columns of both Excel files
        print("\nExcel 1 - Preview (First 5 rows):")
        print(df1.head())
        print("\nExcel 2 - Preview (First 5 rows):")
        print(df2.head())

        # Ensure both Excel files have required columns
        required_columns_excel1 = ['Description', 'Material', 'Rate', 'C/kg']
        required_columns_excel2 = ['Description', 'Material', 'Rate', 'C/kg']

        for col in required_columns_excel1:
            if col not in df1.columns:
                raise ValueError(f"Excel 1 must contain the column '{col}'.")
        for col in required_columns_excel2[:2]:  # Only 'Description' and 'Material' are mandatory in Excel 2
            if col not in df2.columns:
                raise ValueError(f"Excel 2 must contain the column '{col}'.")

        # Fill missing values in Excel 2 with corresponding values from Excel 1
        df2_filled = df2.copy()
        for index, row in df2.iterrows():
            if pd.isnull(row.get('Rate')) or pd.isnull(row.get('C/kg')):
                match = df1[(df1['Description'] == row['Description']) & 
                            (df1['Material'] == row['Material'])]
                if not match.empty:
                    if pd.isnull(row.get('Rate')) and 'Rate' in match.columns:
                        df2_filled.at[index, 'Rate'] = match.iloc[0]['Rate']
                    if pd.isnull(row.get('C/kg')) and 'C/kg' in match.columns:
                        df2_filled.at[index, 'C/kg'] = match.iloc[0]['C/kg']

        # Debugging: Check for missing data
        print("\nUpdated Excel 2 - Preview (First 5 rows):")
        print(df2_filled.head())

        # Check if there are still missing values
        missing_data = df2_filled[df2_filled['Rate'].isnull() | df2_filled['C/kg'].isnull()]
        missing = not missing_data.empty

        # Generate file paths for the output files
        output_file_name = os.path.basename(excel2_path).replace(".xlsx", "_updated.xlsx")
        output_path = os.path.join(output_dir, output_file_name)
        missing_data_file_name = os.path.basename(excel2_path).replace(".xlsx", "_missing_data.xlsx")
        missing_data_path = os.path.join(output_dir, missing_data_file_name)

        # Save the updated Excel 2
        df2_filled.to_excel(output_path, index=False)
        print(f"\nUpdated Excel 2 saved to: {output_path}")

        # If there are missing values, save those rows to a separate file
        if missing:
            missing_data.to_excel(missing_data_path, index=False)
            print(f"Rows with missing data saved to: {missing_data_path}")

        return output_path, missing_data_path if missing else None

    except Exception as e:
        print(f"Error in process_excel: {str(e)}")
        return None, str(e)

if __name__ == "__main__":
    try:
        # Validate command line arguments
        if len(sys.argv) < 3:
            raise ValueError("Usage: python process_excel.py <Excel 1 Path> <Excel 2 Path>")

        # File paths from command line arguments
        excel1_path = sys.argv[1]
        excel2_path = sys.argv[2]

        # Optional: Output directory
        output_dir = "processed"
        if len(sys.argv) == 4:
            output_dir = sys.argv[3]

        # Call the process_excel function
        output_path, missing_data_path = process_excel(excel1_path, excel2_path, output_dir)

        # Display results
        if output_path:
            print(f"\nFiltered file saved to: {output_path}")
            if missing_data_path:
                print(f"Some data is missing. Missing data saved to: {missing_data_path}")
                print("Please upload another Excel1 file to complete the process.")
            else:
                print("Excel 2 is fully populated.")
        else:
            print("Processing failed. Please check the error messages above.")

    except Exception as e:
        print(f"Error: {str(e)}")
