import pandas as pd
import sys
import os

def aggregate_excel_sheets(folder_path, sheet_name):
    # Create an empty DataFrame to store combined data
    combined_df = pd.DataFrame()

    # Iterate through all files in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            try:
                # Read the specified sheet from the current workbook
                sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)

                # Add a new column with the title of the document (filename)
                sheet_df['Source Document'] = filename

                # Append the current DataFrame to the combined DataFrame
                combined_df = pd.concat([combined_df, sheet_df], ignore_index=True)
                print(f"Processed: {filename}")

            except ValueError:
                print(f"Warning: '{sheet_name}' not found in '{filename}'. Skipping.")

            except Exception as e:
                print(f"Error processing '{filename}': {e}")

    # Save the combined DataFrame to a new Excel file
    if not combined_df.empty:
        output_file = os.path.join(folder_path, f'combined_{sheet_name}.xlsx')
        combined_df.to_excel(output_file, index=False)
        print(f"Combined data saved to '{output_file}'")
    else:
        print("No data was aggregated.")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python aggregate_excel_sheets.py <folder_path> <sheet_name>")
    else:
        folder_path = sys.argv[1]
        sheet_name = sys.argv[2]
        aggregate_excel_sheets(folder_path, sheet_name)
