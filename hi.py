import pandas as pd
import xlwings as xw


def read_sf1(sf1_path):
    """
    Reads and extracts relevant data from SF1, organizes it alphabetically by name.
    """
    print(f"Reading SF1 data from {sf1_path}...")
    # Load the data from SF1
    df = pd.read_excel(sf1_path)

    # Extract and split NAME column
    if "NAME" in df.columns:
        name_split = df["NAME"].str.extract(
            r'(?P<Last Name>[\w\-\']+),\s*(?P<First Name>[\w\-\']+)\s*(?P<Name Extension>[A-Z]\.)?\s*(?P<Middle Name>[\w\-\']+)?'
        )
        df = pd.concat([df, name_split], axis=1)
    else:
        raise KeyError("'NAME' column not found in SF1. Please check the file format.")

    # Organize data alphabetically by Last Name, then First Name
    df = df.sort_values(by=["Last Name", "First Name"], ignore_index=True)

    # Filter male and female rows based on the specified row ranges
    male_data = df.iloc[:40]  # Row 11-50 for males
    female_data = df.iloc[40:]  # Row 52-91 for females

    return male_data, female_data


def write_to_sf5(sf5_path, male_data, female_data, sf5_type):
    """
    Writes the organized data into SF5A or SF5B using xlwings.
    """
    print(f"Writing data to {sf5_path} ({sf5_type})...")
    app = xw.App(visible=False)
    wb = app.books.open(sf5_path)
    sheet = wb.sheets[0]  # Assuming single sheet, adjust as necessary

    # Define row ranges and columns for each SF5 type
    if sf5_type == "SF5A":
        male_start_row, female_start_row = 13, 45
        lrn_col, name_cols = "C", ["D", "E", "F", "G", "H", "I"]
    elif sf5_type == "SF5B":
        male_start_row, female_start_row = 15, 45
        lrn_col, name_cols = "B", ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"]
    else:
        raise ValueError("Invalid SF5 type. Must be 'SF5A' or 'SF5B'.")

    # Write male data
    current_row = male_start_row
    for _, row in male_data.iterrows():
        # Write LRN
        sheet.range(f"{lrn_col}{current_row}").value = row["LRN"]
        # Write Name (merge across specified columns)
        full_name = f"{row['Last Name']}, {row['First Name']} {row['Name Extension'] or ''} {row['Middle Name'] or ''}"
        sheet.range(f"{name_cols[0]}{current_row}:{name_cols[-1]}{current_row}").value = full_name.strip()
        current_row += 1

    # Write female data
    current_row = female_start_row
    for _, row in female_data.iterrows():
        # Write LRN
        sheet.range(f"{lrn_col}{current_row}").value = row["LRN"]
        # Write Name (merge across specified columns)
        full_name = f"{row['Last Name']}, {row['First Name']} {row['Name Extension'] or ''} {row['Middle Name'] or ''}"
        sheet.range(f"{name_cols[0]}{current_row}:{name_cols[-1]}{current_row}").value = full_name.strip()
        current_row += 1

    # Save and close
    wb.save()
    wb.close()
    app.quit()
    print(f"Data successfully written to {sf5_path} ({sf5_type}).")


def main():
    # File paths
    sf1_path = "SF1.xlsx"  # Replace with your actual SF1 file path
    sf5a_path = "SF5A.xlsx"  # Replace with your actual SF5A file path
    sf5b_path = "SF5B.xlsx"  # Replace with your actual SF5B file path

    # Step 1: Read and organize data from SF1
    male_data, female_data = read_sf1(sf1_path)

    # Step 2: Write the data to SF5A
    write_to_sf5(sf5a_path, male_data, female_data, "SF5A")

    # Step 3: Write the data to SF5B
    write_to_sf5(sf5b_path, male_data, female_data, "SF5B")


if __name__ == "__main__":
    main()
