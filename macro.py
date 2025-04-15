import xlwings as xw
import os
import shutil
import sqlite3
import time
import pandas as pd
import logging
from datetime import datetime

# Setup logging
logging.basicConfig(
    filename='excel_auto_transfer.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class AutoTransferSystem:
    def __init__(self):
        self.db_path = 'student_records.db'
        self.base_folder = 'SF9SF10'
        self.sf9_folder = os.path.join(self.base_folder, 'SF9')
        self.sf10_folder = os.path.join(self.base_folder, 'SF10')
        self.mfq_paths = ['MFQ1.xlsx', 'MFQ2.xlsx', 'MFQ3.xlsx', 'MFQ4.xlsx']
        
        # Create necessary folders
        self.create_folders()
        
        # Create or connect to database
        self.setup_database()
        
        # Setup Excel application
        self.app = None

    def create_folders(self):
        """Create necessary folders for SF9 and SF10 files"""
        os.makedirs(self.sf9_folder, exist_ok=True)
        os.makedirs(self.sf10_folder, exist_ok=True)

    def setup_database(self):
        """Create or connect to SQLite database"""
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        cursor = conn.cursor()
        
        # Create tables for student info and grades
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            lrn TEXT PRIMARY KEY,
            name TEXT,
            section TEXT,
            grade_level TEXT,
            school_id TEXT,
            school_name TEXT,
            school_year TEXT,
            adviser TEXT,
            gender TEXT,
            birth_date TEXT,
            age TEXT,
            mother_tongue TEXT,
            ip_community TEXT,
            father_name TEXT,
            mother_name TEXT,
            guardian_name TEXT,
            contact_number TEXT
        )
        ''')
        
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS grades (
            lrn TEXT,
            subject_idx INTEGER,
            quarter INTEGER,
            grade REAL,
            PRIMARY KEY (lrn, subject_idx, quarter),
            FOREIGN KEY (lrn) REFERENCES students(lrn)
        )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grades_lrn ON grades(lrn)')
        
        conn.commit()
        conn.close()
        logging.info("Database setup completed")

    def install_excel_macros(self):
        """Install macros in MFQ files to trigger data transfer on change"""
        self.app = xw.App(visible=False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        
        try:
            for path in self.mfq_paths:
                self.install_macro_in_file(path)
            
            logging.info("Macros installed in all MFQ files")
        finally:
            self.app.quit()
            self.app = None

    def install_macro_in_file(self, file_path):
        """Install auto-transfer macro in a specific Excel file"""
        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path}")
            return
        
        try:
            # Open the workbook
            wb = self.app.books.open(file_path)
            
            # Check if VBA module exists, add if not
            try:
                vba_module = wb.vba_modules.add("AutoTransferModule")
            except:
                vba_module = wb.vba_modules["AutoTransferModule"]
            
            # VBA code for auto-transfer
            vba_code = """
            Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
                ' Check if changes are in the data area
                If Target.Row >= 6 And Target.Row <= 100 Then
                    If (Target.Column >= 4 And Target.Column <= 12) Or Target.Column = 1 Then
                        ' Save the workbook
                        ThisWorkbook.Save
                        
                        ' Create a trigger file to notify Python script
                        Dim filePath As String
                        filePath = ThisWorkbook.Path & "\\data_changed.trigger"
                        
                        ' Write file info to trigger
                        Dim fileNum As Integer
                        fileNum = FreeFile
                        Open filePath For Output As fileNum
                        Print #fileNum, ThisWorkbook.Name
                        Print #fileNum, Target.Worksheet.Name
                        Print #fileNum, Target.Address
                        Print #fileNum, Now
                        Close fileNum
                    End If
                End If
            End Sub
            """
            
            # Set the VBA code
            vba_module.code = vba_code
            
            # Save the workbook
            wb.save()
            wb.close()
            logging.info(f"Macro installed in {file_path}")
            
        except Exception as e:
            logging.error(f"Error installing macro in {file_path}: {str(e)}")

    def start_monitoring(self):
        """Start monitoring for data changes and process transfers"""
        trigger_path = os.path.join(os.path.dirname(self.mfq_paths[0]), "data_changed.trigger")
        logging.info("Starting monitoring for data changes...")
        
        try:
            while True:
                if os.path.exists(trigger_path):
                    try:
                        # Read the trigger file
                        with open(trigger_path, 'r') as f:
                            workbook_name = f.readline().strip()
                            sheet_name = f.readline().strip()
                            cell_address = f.readline().strip()
                            timestamp = f.readline().strip()
                        
                        # Remove the trigger file
                        os.remove(trigger_path)
                        
                        # Process the data change
                        logging.info(f"Data change detected in {workbook_name}, {sheet_name}, {cell_address}")
                        self.process_data_change(workbook_name)
                        
                    except Exception as e:
                        logging.error(f"Error processing trigger: {str(e)}")
                        if os.path.exists(trigger_path):
                            os.remove(trigger_path)
                
                # Sleep to avoid high CPU usage
                time.sleep(2)
                
        except KeyboardInterrupt:
            logging.info("Monitoring stopped by user")
            print("Monitoring stopped.")

    def process_data_change(self, workbook_name):
        """Process the data change in the specified workbook"""
        try:
            # Extract the quarter from the workbook name
            if workbook_name.startswith("MFQ"):
                quarter = int(workbook_name[3:4])
            else:
                quarter = 1  # Default to first quarter if not recognized
            
            # Update database from all MFQ files for consistency
            self.update_database_from_excel()
            
            # Get affected students and process their SF files
            affected_students = self.get_affected_students(quarter)
            if affected_students:
                self.update_sf_files(affected_students)
                
            logging.info(f"Processed changes for {len(affected_students)} students")
            
        except Exception as e:
            logging.error(f"Error processing data change: {str(e)}")

    def update_database_from_excel(self):
        """Update the database with current data from Excel files"""
        conn = sqlite3.connect(self.db_path)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        try:
            # Open all MFQ workbooks
            mfq_workbooks = [app.books.open(path) for path in self.mfq_paths]
            
            # Process student information from MFQ1
            wb_mfq1 = mfq_workbooks[0]
            sheet = wb_mfq1.sheets[0]
            
            student_ranges = list(range(6, 50)) + list(range(52, 101))
            
            # Get all student LRNs to check which have data
            student_lrns = []
            for student_row in student_ranges:
                lrn = sheet.range(f'A{student_row}').value
                if lrn:
                    student_lrns.append(lrn)
            
            # Clear existing data for these students to avoid duplicates
            cursor = conn.cursor()
            placeholders = ','.join(['?' for _ in student_lrns])
            cursor.execute(f"DELETE FROM students WHERE lrn IN ({placeholders})", student_lrns)
            cursor.execute(f"DELETE FROM grades WHERE lrn IN ({placeholders})", student_lrns)
            
            # Prepare batch data for students and grades
            student_data = []
            grades_data = []
            
            adviser_value = sheet.range('R28').value
            
            # Batch process student information
            for student_row in student_ranges:
                lrn = sheet.range(f'A{student_row}').value
                if not lrn:
                    continue
                
                # Extract student info
                student_info = (
                    lrn,
                    sheet.range(f'B{student_row}').value,
                    sheet.range(f'Q{student_row}').value,
                    sheet.range(f'P{student_row}').value,
                    sheet.range(f'S{student_row}').value,
                    sheet.range(f'R{student_row}').value,
                    sheet.range(f'T{student_row}').value,
                    adviser_value,
                    'Male' if student_row < 51 else 'Female',
                    sheet.range(f'Q{student_row}').value,
                    sheet.range(f'F{student_row}').value,
                    sheet.range(f'AW{student_row}').value,
                    sheet.range(f'AX{student_row}').value,
                    sheet.range(f'AY{student_row}').value,
                    sheet.range(f'BA{student_row}').value,
                    sheet.range(f'BC{student_row}').value,
                    sheet.range(f'BB{student_row}').value
                )
                
                student_data.append(student_info)
                
                # Extract grades for all quarters
                rows = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
                
                for q_idx, wb in enumerate(mfq_workbooks, 1):
                    for subject_idx, row in enumerate(rows):
                        if subject_idx == 8 and q_idx < 3:
                            continue
                        
                        grade = wb.sheets[0].range(f'{row}{student_row}').value
                        if grade is not None:
                            grades_data.append((lrn, subject_idx, q_idx, grade))
            
            # Insert data in a single transaction
            with conn:
                conn.executemany(
                    'INSERT OR REPLACE INTO students VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                    student_data
                )
                
                conn.executemany(
                    'INSERT OR REPLACE INTO grades VALUES (?, ?, ?, ?)',
                    grades_data
                )
            
            logging.info(f"Updated database with {len(student_data)} students and {len(grades_data)} grade records")
            
        except Exception as e:
            logging.error(f"Error updating database: {str(e)}")
        finally:
            # Close workbooks and Excel application
            for wb in mfq_workbooks:
                wb.close()
            app.quit()
            conn.close()

    def get_affected_students(self, quarter):
        """Get students affected by changes in the specified quarter"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Get all students with grades in the specified quarter
        cursor.execute("""
            SELECT DISTINCT s.* 
            FROM students s
            JOIN grades g ON s.lrn = g.lrn
            WHERE g.quarter = ?
        """, (quarter,))
        
        students = cursor.fetchall()
        conn.close()
        
        return students

    def update_sf_files(self, students):
        """Update SF9 and SF10 files for the specified students"""
        conn = sqlite3.connect(self.db_path)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        
        try:
            for student in students:
                lrn = student[0]
                
                # Get student's grades
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM grades WHERE lrn = ?", (lrn,))
                grades = cursor.fetchall()
                
                # Copy template files if they don't exist
                sf9_path = os.path.join(self.sf9_folder, f"{lrn}.xlsb")
                sf10_path = os.path.join(self.sf10_folder, f"{lrn}.xlsx")
                
                if not os.path.exists(sf9_path):
                    shutil.copy('SF9.xlsb', sf9_path)
                
                if not os.path.exists(sf10_path):
                    shutil.copy('sf10.xlsx', sf10_path)
                
                # Open the workbooks
                wb_sf9 = app.books.open(sf9_path)
                wb_sf10 = app.books.open(sf10_path)
                
                try:
                    # Update front page information
                    self.update_front_page(student, wb_sf9, wb_sf10)
                    
                    # Update grades
                    self.update_grades(grades, wb_sf9, wb_sf10)
                    
                    # Save the workbooks
                    wb_sf9.save()
                    wb_sf10.save()
                    
                    logging.info(f"Updated SF files for student {lrn}")
                    
                finally:
                    wb_sf9.close()
                    wb_sf10.close()
                
        except Exception as e:
            logging.error(f"Error updating SF files: {str(e)}")
        finally:
            app.quit()
            conn.close()

    def update_front_page(self, student, wb_sf9, wb_sf10):
        """Update front page data in SF9 and SF10 files"""
        front_sf9 = wb_sf9.sheets['FRONT']
        front_sf10 = wb_sf10.sheets['FRONT']
        
        # Unpack student data
        (lrn, name, section, grade_level, school_id, school_name, school_year, 
         adviser, gender, birth_date, age, mother_tongue, ip_community, 
         father_name, mother_name, guardian_name, contact_number) = student
        
        # SF9 Front page updates
        sf9_data = {
            'Q22': name,
            'T3': lrn,
            'Q26': section,
            'R29': school_name,
            'P40': grade_level,
            'S40': school_id,
            'R28': adviser,
            'Q24': contact_number
        }
        
        # SF10 Front page updates
        sf10_data = {
            'AS66': section,
            'G68': school_name,
            'A92': school_id,
            'BA66': adviser,
            'F8': mother_tongue,
            'Y8': ip_community,
            'AZ8': father_name,
            'C9': lrn,
            'AA9': guardian_name,
            'AN9': mother_name
        }
        
        # Apply updates
        for cell, value in sf9_data.items():
            front_sf9.range(cell).value = value
        
        for cell, value in sf10_data.items():
            front_sf10.range(cell).value = value

    def update_grades(self, grades, wb_sf9, wb_sf10):
        """Update grades in SF9 and SF10 files"""
        back_sf9 = wb_sf9.sheets['BACK']
        front_sf10 = wb_sf10.sheets['FRONT']
        
        # Prepare quarter mappings
        quarter_mappings = {
            1: {'sf9_col': 'C', 'sf9_start_row': 7, 'sf10_col': 'AT', 'sf10_start_row': 31},
            2: {'sf9_col': 'D', 'sf9_start_row': 7, 'sf10_col': 'AY', 'sf10_start_row': 31},
            3: {'sf9_col': 'C', 'sf9_start_row': 23, 'sf10_col': 'AT', 'sf10_start_row': 74},
            4: {'sf9_col': 'D', 'sf9_start_row': 23, 'sf10_col': 'AY', 'sf10_start_row': 74}
        }
        
        # Prepare updates
        sf9_updates = {}
        sf10_updates = {}
        
        for grade_info in grades:
            lrn, subject_idx, quarter, grade = grade_info
            
            # Skip if grade is None
            if grade is None:
                continue
            
            # Get quarter mapping
            mapping = quarter_mappings[quarter]
            
            # Calculate row offsets
            sf9_row = mapping['sf9_start_row'] + subject_idx
            # Adjust rows for specific subjects
            if subject_idx in [3, 4, 5]:
                sf9_row += 1
            elif subject_idx in [6, 7]:
                sf9_row += 2
            elif subject_idx == 8:
                sf9_row += 2
            
            # Prepare updates
            sf9_cell = f"{mapping['sf9_col']}{sf9_row}"
            sf10_cell = f"{mapping['sf10_col']}{mapping['sf10_start_row'] + subject_idx}"
            
            sf9_updates[sf9_cell] = grade
            sf10_updates[sf10_cell] = grade
        
        # Apply updates
        for cell, value in sf9_updates.items():
            back_sf9.range(cell).value = value
        
        for cell, value in sf10_updates.items():
            front_sf10.range(cell).value = value

def main():
    print("Starting Excel Auto-Transfer System")
    print("==================================")
    
    system = AutoTransferSystem()
    
    print("\nMenu Options:")
    print("1. Install macros in MFQ files")
    print("2. Start monitoring for data changes")
    print("3. Force update all SF files")
    print("4. Exit")
    
    while True:
        try:
            choice = input("\nEnter your choice (1-4): ")
            
            if choice == '1':
                print("Installing macros in MFQ files...")
                system.install_excel_macros()
                print("Macros installed successfully.")
                
            elif choice == '2':
                print("Starting monitoring for data changes...")
                print("Press Ctrl+C to stop monitoring.")
                system.start_monitoring()
                
            elif choice == '3':
                print("Forcing update of all SF files...")
                system.update_database_from_excel()
                
                # Get all students
                conn = sqlite3.connect(system.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM students")
                students = cursor.fetchall()
                conn.close()
                
                # Update SF files for all students
                print(f"Updating SF files for {len(students)} students...")
                system.update_sf_files(students)
                print("All SF files updated successfully.")
                
            elif choice == '4':
                print("Exiting...")
                break
                
            else:
                print("Invalid choice. Please try again.")
                
        except Exception as e:
            print(f"Error: {str(e)}")
            logging.error(f"Error in main menu: {str(e)}")

if __name__ == "__main__":
    main()