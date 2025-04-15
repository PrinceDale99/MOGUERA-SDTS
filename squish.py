import xlwings as xw

def extract_name_parts(full_name):
    # Handle None or empty values
    if not full_name or isinstance(full_name, list):  # Added list check
        return None, None, None
        
    # Convert to string if not already
    full_name = str(full_name).strip()
    if not full_name:
        return None, None, None
        
    parts = full_name.split(",")
    if len(parts) < 2:
        print(f"Invalid format for name: {full_name}. Skipping...")
        return None, None, None
        
    last_name = parts[0].strip()
    remaining_part = parts[1].strip()
    name_parts = remaining_part.split()
    
    if not name_parts:  # If no first name/middle initial
        return None, None, None
        
    # Handle case where there's no middle initial
    if len(name_parts) == 1:
        return last_name, name_parts[0], ""
    
    # Normal case with middle initial
    middle_initial = name_parts[-1]
    first_name = " ".join(name_parts[:-1])
    return last_name, first_name, middle_initial

def transfer_student_details():
    try:
        # Open both workbooks
        app = xw.App(visible=False)
        sf_wb = app.books.open('sf1.xlsx')
        mfq_wb = app.books.open('MFQ1.xlsx')
        
        # Get active sheets
        sf_sheet = sf_wb.sheets.active
        mfq_sheet = mfq_wb.sheets.active
        
        def process_students(start_row, end_row, target_start_row):
            for row in range(start_row, end_row + 1):
                try:
                    # Read source data (sf1.xlsx)
                    lrn = sf_sheet.range(f'B{row}').value
                    
                    # Check if row is empty
                    if not lrn:
                        continue
                        
                    full_name = sf_sheet.range(f'C{row}').value  # Changed to only read column C
                    gender = sf_sheet.range(f'G{row}').value
                    age = sf_sheet.range(f'H{row}').value  # Changed to only read column H
                    birth_date = sf_sheet.range(f'J{row}').value  # Changed to only read column J
                    
                    # Skip if no name is found
                    if not full_name:
                        print(f"No name found in row {row}. Skipping...")
                        continue
                    
                    # Extract name parts
                    name_result = extract_name_parts(full_name)
                    
                    # Skip if name parsing failed
                    if name_result[0] is None:
                        continue
                        
                    last_name, first_name, middle_initial = name_result
                    
                    # Calculate target row
                    target_row = target_start_row + (row - start_row)
                    
                    # Write to MFQ1.xlsx
                    mfq_sheet.range(f'A{target_row}').value = lrn
                    mfq_sheet.range(f'B{target_row}').value = full_name
                    mfq_sheet.range(f'BA{target_row}').value = gender
                    mfq_sheet.range(f'BC{target_row}').value = age
                    mfq_sheet.range(f'BB{target_row}').value = birth_date
                    mfq_sheet.range(f'AW{target_row}').value = last_name
                    mfq_sheet.range(f'AX{target_row}').value = first_name
                    mfq_sheet.range(f'AY{target_row}').value = middle_initial
                    
                    print(f"Processed row {row}: {full_name} -> {last_name}, {first_name}, {middle_initial}")
                    
                except Exception as e:
                    print(f"Error processing row {row}: {str(e)}")
                    continue

        print("Processing male students...")
        process_students(11, 50, 6)
        
        print("\nProcessing female students...")
        process_students(52, 91, 52)
        
        # Save changes
        mfq_wb.save()
        
        print("\nTransfer completed successfully!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        
    finally:
        # Cleanup
        sf_wb.close()
        mfq_wb.close()
        app.quit()

if __name__ == "__main__":
    transfer_student_details()