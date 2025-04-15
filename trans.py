import xlwings as xw
from concurrent.futures import ThreadPoolExecutor
import time

def extract_name_parts(full_name):
    """Extract last name, first name, and middle initial from a full name."""
    # Handle None or empty values
    if not full_name or isinstance(full_name, list):
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

def process_students(sf_sheet, mfq_sheet, start_row, end_row, target_start_row):
    """Process student data from source to target worksheet."""
    # Read all data at once for better performance
    source_range = sf_sheet.range(f'B{start_row}:J{end_row}')
    source_data = source_range.value
    
    # Prepare data for bulk writing
    target_data = {}
    for columns in ['A', 'B', 'BA', 'BC', 'BB', 'AW', 'AX', 'AY']:
        target_data[columns] = []
    
    for row_data in source_data:
        # Skip if row is None or first element (LRN) is None
        if not row_data or row_data[0] is None:
            continue
            
        lrn = row_data[0]  # Column B
        full_name = row_data[1]  # Column C
        
        # Skip if name is missing
        if not full_name:
            continue
            
        gender = row_data[5] if len(row_data) > 5 else None  # Column G
        age = row_data[6] if len(row_data) > 6 else None  # Column H
        birth_date = row_data[8] if len(row_data) > 8 else None  # Column J
        
        # Extract name parts
        name_result = extract_name_parts(full_name)
        
        # Skip if name parsing failed
        if name_result[0] is None:
            continue
            
        last_name, first_name, middle_initial = name_result
        
        # Add to data for bulk writing
        target_data['A'].append(lrn)
        target_data['B'].append(full_name)
        target_data['BA'].append(gender)
        target_data['BC'].append(age)
        target_data['BB'].append(birth_date)
        target_data['AW'].append(last_name)
        target_data['AX'].append(first_name)
        target_data['AY'].append(middle_initial)
        
        print(f"Processed: {full_name} -> {last_name}, {first_name}, {middle_initial}")
    
    # Bulk write data to target sheet
    for column, values in target_data.items():
        if values:  # Only write if there's data
            target_rows = len(values)
            write_range = f'{column}{target_start_row}:{column}{target_start_row + target_rows - 1}'
            mfq_sheet.range(write_range).options(transpose=True).value = values
    
    return len(target_data['A']), target_data  # Return number of rows processed and data

def transfer_student_details():
    """Transfer student details from sf1.xlsx to MFQ1.xlsx (squish.py functionality)."""
    app = None
    sf_wb = None
    mfq_wb = None
    
    try:
        # Open both workbooks with a single Excel instance
        app = xw.App(visible=False)
        app.display_alerts = False  # Suppress any Excel alerts
        sf_wb = app.books.open('sf1.xlsx')
        mfq_wb = app.books.open('MFQ1.xlsx')
        
        # Get active sheets
        sf_sheet = sf_wb.sheets.active
        mfq_sheet = mfq_wb.sheets.active
        
        print("Processing male students...")
        male_count, male_data = process_students(sf_sheet, mfq_sheet, 11, 50, 6)
        
        print("\nProcessing female students...")
        female_count, female_data = process_students(sf_sheet, mfq_sheet, 52, 91, 52)
        
        # Combine data for direct transfer to other workbooks
        combined_data = {
            'lrns': male_data['A'] + female_data['A'],
            'names': male_data['B'] + female_data['B']
        }
        
        # Save changes
        mfq_wb.save()
        
        print(f"\nTransfer completed successfully! Processed {male_count} male and {female_count} female students.")
        return True, combined_data
        
    except Exception as e:
        print(f"An error occurred in transfer_student_details: {str(e)}")
        return False, None
        
    finally:
        # Cleanup
        try:
            if sf_wb is not None:
                sf_wb.close()
            if mfq_wb is not None:
                mfq_wb.close()
            if app is not None:
                app.quit()
        except Exception as e:
            print(f"Error during cleanup: {str(e)}")

def process_single_file(app, target_file, lrns, names):
    """Process a single target workbook using the provided Excel app instance."""
    target_wb = None
    
    try:
        target_wb = app.books.open(target_file, update_links=False)
        target_sheet = target_wb.sheets[0]
        
        # Determine how many male and female entries we have
        male_count = min(len(lrns[:44]), 44)
        female_count = len(lrns[44:])
        
        # Bulk write data for better performance
        # Write male students (rows 6-49)
        if male_count > 0:
            target_sheet.range(f'A6:A{5+male_count}').options(transpose=True).value = lrns[:male_count]
            target_sheet.range(f'B6:B{5+male_count}').options(transpose=True).value = names[:male_count]
        
        # Write female students (rows 52-100)
        if female_count > 0:
            target_sheet.range(f'A52:A{51+female_count}').options(transpose=True).value = lrns[44:44+female_count]
            target_sheet.range(f'B52:B{51+female_count}').options(transpose=True).value = names[44:44+female_count]
        
        # Save workbook
        target_wb.save()
        print(f"Data successfully transferred to {target_file}")
        return True
        
    except Exception as e:
        print(f"Error processing {target_file}: {str(e)}")
        return False
        
    finally:
        # Close workbook but not the app
        try:
            if target_wb is not None:
                target_wb.close()
        except Exception as e:
            print(f"Error closing {target_file}: {str(e)}")

def transfer_data(data=None):
    """Transfer data from MFQ1.xlsx to other MFQ files (brock.py functionality)."""
    app = None
    source_wb = None
    
    try:
        # If data is already provided from transfer_student_details, use it
        if data:
            lrns = data['lrns']
            names = data['names']
        else:
            # Otherwise, open the source workbook (MFQ1)
            app = xw.App(visible=False)
            app.display_alerts = False
            source_wb = app.books.open('MFQ1.xlsx')
            source_sheet = source_wb.sheets[0]
            
            # Read data in bulk for better performance
            lrns_male = source_sheet.range('A6:A49').value
            names_male = source_sheet.range('B6:B49').value
            lrns_female = source_sheet.range('A52:A100').value
            names_female = source_sheet.range('B52:B100').value
            
            # Filter out None values
            lrns_male = [lrn for lrn in lrns_male if lrn is not None]
            names_male = [name for name in names_male if name is not None]
            lrns_female = [lrn for lrn in lrns_female if lrn is not None]
            names_female = [name for name in names_female if name is not None]
            
            # Combine lists
            lrns = lrns_male + lrns_female
            names = names_male + names_female
            
            # Close source workbook but keep app open
            if source_wb is not None:
                source_wb.close()
            source_wb = None
        
        # Create a single Excel instance if not already provided
        if app is None:
            app = xw.App(visible=False)
            app.display_alerts = False
        
        # Target files to update using a single Excel instance
        target_files = ['MFQ2.xlsx', 'MFQ3.xlsx', 'MFQ4.xlsx']
        
        print("Starting data transfer to target files...")
        
        # Process each file sequentially with the same app instance
        for target_file in target_files:
            try:
                success = process_single_file(app, target_file, lrns, names)
                if not success:
                    print(f"Failed to process {target_file}")
                # Add a short delay between file operations
                time.sleep(0.5)
            except Exception as e:
                print(f"Exception occurred while processing {target_file}: {str(e)}")
        
        print("Data transfer to all files completed!")
        return True
        
    except Exception as e:
        print(f"An error occurred in transfer_data: {str(e)}")
        return False
        
    finally:
        # Clean up the Excel app
        try:
            if source_wb is not None:
                source_wb.close()
            if app is not None:
                app.quit()
        except Exception as e:
            print(f"Error during cleanup in transfer_data: {str(e)}")

def main():
    """Main function to run processes with optimized workflow."""
    print("Step 1: Transferring student details from sf1.xlsx to MFQ1.xlsx...")
    success, combined_data = transfer_student_details()
    
    if success:
        print("\nStep 2: Transferring data to MFQ2, MFQ3, and MFQ4 using a single Excel instance...")
        # Use the data already collected in Step 1 directly, avoiding re-reading from MFQ1
        transfer_data(combined_data)
    else:
        print("Aborting process due to failure in Step 1.")

if __name__ == "__main__":
    main()