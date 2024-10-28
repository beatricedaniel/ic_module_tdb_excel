import pandas as pd
import glob
import os

def list_csv_files(directory: str):
    """Return list of all .csv files in a given directory."""
    return glob.glob(os.path.join(directory, "*.csv"))

def find_files(pattern: str):
    """Return list of files matching a pattern."""
    return glob.glob(pattern)

def process_csv_files_to_sheets(directory: str, output_file: str):
    """Process all CSV files in a directory and save them as sheets in an .xlsx file."""
    # Find all CSV files in the directory
    csv_files = list_csv_files(directory)
    print(csv_files)
    if not csv_files:
        print("No CSV files found in the specified directory.")
        return

    # Open Excel writer to write sheets
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        if not csv_files:
            # Add an empty sheet if there are no CSV files, to avoid the empty file error
            pd.DataFrame().to_excel(writer, sheet_name="Empty")
        else:
            for csv_file in csv_files:
                # Extract file name to use as sheet name
                sheet_name = os.path.splitext(os.path.basename(csv_file))[0][:30]  # Limit to 30 chars
                # Load CSV as DataFrame
                df = pd.read_csv(csv_file, sep=";")
                # Save each DataFrame as a separate sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"All CSV files from {directory} have been saved to {output_file}")

def create_crosstab_and_add_to_excel(file_path: str, last_column: str, writer, sheet_name: str):
    """Create a crosstab from a specific CSV file and add it to an Excel file."""
    # Load CSV as DataFrame
    df = pd.read_csv(file_path, sep=";")
    print(df)
    
    # Create crosstab
    crosstab = pd.crosstab(
        index=[df["Région"], df["Statut juridique"]],
        columns=[df["Thème Décision"], df["Sous-thème Décision"]],
        values=df[last_column],
        aggfunc="sum",
        margins=False
    ).sort_index(level=["Région", "Statut juridique"])

    # Save crosstab to the Excel writer
    crosstab.to_excel(writer, sheet_name=sheet_name)

def main():
    # Directory containing general CSV files
    general_csv_directory = "../../../2024_10_output/tdb/"  # Replace with the actual directory path
    # Output .xlsx file name
    output_xlsx = "tdb_inspection_controle.xlsx"
    
    # Step 1 & 2: Process general CSV files into sheets
    process_csv_files_to_sheets(general_csv_directory, output_xlsx)

    # Specific CSV files with distinct columns for crosstab
    specific_files_info = [
        ("../../../2024_10_output/tdb/TDB_INJONCTION_*.csv", "Injonctions", "tcd_Injonctions"),
        ("../../../2024_10_output/tdb/TDB_PRESCRIPTION_*.csv", "Prescriptions", "tcd_Prescriptions"),
        ("../../../2024_10_output/tdb/TDB_INJONCTIONS_PRESCRIPTIONS_*.csv", "Injonctions + prescriptions", "tcd_InjonctionsPrescriptions")
    ]

    # Step 3-7: Process specific CSVs, create crosstabs, and add them to the same .xlsx file
    with pd.ExcelWriter(output_xlsx, engine="openpyxl", mode="a") as writer:
        for file_pattern, last_column, sheet_name in specific_files_info:
            # Find the latest file matching the pattern
            specific_csv_files = find_files(file_pattern)
            if specific_csv_files:
                print(file_pattern)
                print(specific_csv_files)
                print(last_column)
                print(sheet_name)
                specific_file = specific_csv_files[-1]  # Use the most recent file based on date order
                # Create crosstab and add to the Excel writer
                create_crosstab_and_add_to_excel(specific_file, last_column, writer, sheet_name)

if __name__ == "__main__":
    main()
