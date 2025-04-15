import xlwings as xw

def transfer_data():
    try:
        # Open the source workbook (MFQ1)
        source_wb = xw.Book('MFQ1.xlsx')
        source_sheet = source_wb.sheets[0]  # Assuming data is in first sheet
        
        # Initialize lists to store data from specific rows
        lrns = []
        names = []
        
        # Read rows 6-49
        for row in range(6, 50):
            lrns.append(source_sheet.range(f'A{row}').value)
            names.append(source_sheet.range(f'B{row}').value)
        
        # Read rows 52-100
        for row in range(52, 101):
            lrns.append(source_sheet.range(f'A{row}').value)
            names.append(source_sheet.range(f'B{row}').value)
        
        # Target files to update
        target_files = ['MFQ2.xlsx', 'MFQ3.xlsx', 'MFQ4.xlsx']
        
        for target_file in target_files:
            # Open target workbook
            target_wb = xw.Book(target_file)
            target_sheet = target_wb.sheets[0]
            
            # Copy data to rows 6-49
            for idx, (lrn, name) in enumerate(zip(lrns[:44], names[:44]), start=6):
                target_sheet.range(f'A{idx}').value = lrn
                target_sheet.range(f'B{idx}').value = name
            
            # Copy data to rows 52-100
            for idx, (lrn, name) in enumerate(zip(lrns[44:], names[44:]), start=52):
                target_sheet.range(f'A{idx}').value = lrn
                target_sheet.range(f'B{idx}').value = name
            
            # Save and close target workbook
            target_wb.save()
            target_wb.close()
            
            print(f"Data successfully transferred to {target_file}")
        
        # Close source workbook
        source_wb.save()
        source_wb.close()
        
        print("Data transfer completed!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    transfer_data()