import xlwings as xw
import os
import shutil
import sqlite3
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import time
import mmap

def create_folders():
    base_folder = 'SF9SF10'
    sf9_folder = os.path.join(base_folder, 'SF9')
    sf10_folder = os.path.join(base_folder, 'SF10')
    
    os.makedirs(sf9_folder, exist_ok=True)
    os.makedirs(sf10_folder, exist_ok=True)
    return sf9_folder, sf10_folder

def create_database():
    """Create SQLite database for storing student data"""
    conn = sqlite3.connect('student_records.db')
    conn.execute("PRAGMA journal_mode=WAL")  # Use WAL mode for better concurrency
    conn.execute("PRAGMA synchronous=NORMAL")  # Reduce synchronous writes for speed
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
    
    # Add indexes for faster lookups
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_grades_lrn ON grades(lrn)')
    
    conn.commit()
    return conn

def load_data_from_excel_to_db(conn, mfq_paths):
    """Extract data from Excel files and store in SQLite database"""
    cursor = conn.cursor()
    
    # Load all MFQ workbooks once to extract data
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    
    try:
        mfq_workbooks = [app.books.open(path) for path in mfq_paths]
        
        # Process student information from MFQ1
        wb_mfq1 = mfq_workbooks[0]
        sheet = wb_mfq1.sheets[0]
        
        student_ranges = list(range(6, 50)) + list(range(52, 101))
        
        # Prepare batch data for students and grades
        student_data = []
        grades_data = []
        
        adviser_value = sheet.range('R28').value  # Get adviser once
        
        # Batch process student information
        for student_row in student_ranges:
            lrn = sheet.range(f'A{student_row}').value
            if not lrn:
                continue
                
            # Extract all student information at once
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
            
            # Extract and insert grades for all quarters
            rows = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
            
            for quarter, wb in enumerate(mfq_workbooks, 1):
                for subject_idx, row in enumerate(rows):
                    if subject_idx == 8 and quarter < 3:
                        continue  # Skip 'L' for quarters 1 and 2
                    
                    grade = wb.sheets[0].range(f'{row}{student_row}').value
                    if grade is not None:
                        grades_data.append((lrn, subject_idx, quarter, grade))
        
        # Insert all student data in a single transaction
        with conn:
            conn.executemany(
                'INSERT OR REPLACE INTO students VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                student_data
            )
            
            # Insert all grades data in the same transaction
            conn.executemany(
                'INSERT OR REPLACE INTO grades VALUES (?, ?, ?, ?)',
                grades_data
            )
    
    finally:
        # Close workbooks and Excel application
        for wb in mfq_workbooks:
            wb.close()
        app.quit()

def copy_template_for_student(lrn, sf9_folder, sf10_folder):
    sf9_path = os.path.join(sf9_folder, f'{lrn}.xlsb')
    sf10_path = os.path.join(sf10_folder, f'{lrn}.xlsx')
    
    if not os.path.exists(sf9_path):
        shutil.copy('SF9.xlsb', sf9_path)
    if not os.path.exists(sf10_path):
        shutil.copy('sf10.xlsx', sf10_path)
    
    return sf9_path, sf10_path

def process_student_batch(batch_lrns, student_dict, grades_dict, sf9_folder, sf10_folder):
    """Process a batch of students in a single Excel instance"""
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    
    try:
        for lrn in batch_lrns:
            sf9_path, sf10_path = copy_template_for_student(lrn, sf9_folder, sf10_folder)
            
            # Get student data
            student = student_dict.get(lrn)
            if not student:
                continue
                
            # Open workbooks
            wb_sf9 = app.books.open(sf9_path)
            wb_sf10 = app.books.open(sf10_path)
            
            try:
                # Process front page data
                process_front_page(student, wb_sf9, wb_sf10)
                
                # Process grades
                student_grades = grades_dict.get(lrn, [])
                process_grades(student_grades, wb_sf9, wb_sf10)
                
                # Save and close workbooks
                wb_sf9.save()
                wb_sf10.save()
            finally:
                wb_sf9.close()
                wb_sf10.close()
    finally:
        app.quit()
    
    return len(batch_lrns)

def process_student_files(student_data, grades_data, sf9_folder, sf10_folder, max_workers=4, batch_size=5):
    """Process SF9 and SF10 files in parallel batches using ThreadPoolExecutor"""
    # Convert student data to dictionary for faster lookup
    student_dict = {student[0]: student for student in student_data}
    
    # Group grades by LRN for faster lookup
    grades_dict = {}
    for grade in grades_data:
        lrn = grade[0]
        if lrn not in grades_dict:
            grades_dict[lrn] = []
        grades_dict[lrn].append(grade)
    
    # Split students into batches
    student_lrns = list(student_dict.keys())
    batches = [student_lrns[i:i+batch_size] for i in range(0, len(student_lrns), batch_size)]
    
    # Process batches in parallel
    total_processed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                process_student_batch, 
                batch, 
                student_dict, 
                grades_dict, 
                sf9_folder, 
                sf10_folder
            ): batch for batch in batches
        }
        
        for future in as_completed(futures):
            batch_processed = future.result()
            total_processed += batch_processed
            print(f"Processed {total_processed}/{len(student_lrns)} students")

def process_front_page(student, wb_sf9, wb_sf10):
    """Process front page data using batch operations"""
    front_sf9 = wb_sf9.sheets['FRONT']
    front_sf10 = wb_sf10.sheets['FRONT']
    
    # Unpack student data
    (lrn, name, section, grade_level, school_id, school_name, school_year, 
     adviser, gender, birth_date, age, mother_tongue, ip_community, 
     father_name, mother_name, guardian_name, contact_number) = student
    
    # SF9 Front page - batch update
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
    
    # SF10 Front page - batch update
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
    
    # Apply all front page updates in one go
    for cell, value in sf9_data.items():
        front_sf9.range(cell).value = value
    
    for cell, value in sf10_data.items():
        front_sf10.range(cell).value = value

def process_grades(student_grades, wb_sf9, wb_sf10):
    """Process grades for all quarters at once using batch operations"""
    back_sf9 = wb_sf9.sheets['BACK']
    front_sf10 = wb_sf10.sheets['FRONT']
    
    # Prepare quarter mappings with fixed destination cells for 3rd and 4th quarters
    quarter_mappings = {
        1: {'sf9_col': 'C', 'sf9_start_row': 7, 'sf10_col': 'AT', 'sf10_start_row': 31},
        2: {'sf9_col': 'D', 'sf9_start_row': 7, 'sf10_col': 'AY', 'sf10_start_row': 31},
        3: {'sf9_col': 'C', 'sf9_start_row': 23, 'sf10_col': 'AT', 'sf10_start_row': 74},
        4: {'sf9_col': 'D', 'sf9_start_row': 23, 'sf10_col': 'AY', 'sf10_start_row': 74}
    }
    
    # Create direct mappings for MFQ3 and MFQ4 fields to correct cells
    # Format: (subject_idx, quarter): sf9_cell
    direct_mappings = {
        (0, 3): 'C23',  # MFQ3 D
        (1, 3): 'C24',  # MFQ3 E
        (2, 3): 'C25',  # MFQ3 F
        (3, 3): 'C26',  # MFQ3 G
        (4, 3): 'C27',  # MFQ3 H
        (5, 3): 'C28',  # MFQ3 I
        (6, 3): 'C30',  # MFQ3 J
        (7, 3): 'C31',  # MFQ3 K
        (8, 3): 'C32',  # MFQ3 L
        (0, 4): 'D23',  # MFQ4 D
        (1, 4): 'D24',  # MFQ4 E
        (2, 4): 'D25',  # MFQ4 F
        (3, 4): 'D26',  # MFQ4 G
        (4, 4): 'D27',  # MFQ4 H
        (5, 4): 'D28',  # MFQ4 I
        (6, 4): 'D30',  # MFQ4 J
        (7, 4): 'D31',  # MFQ4 K
        (8, 4): 'D32',  # MFQ4 L
    }
    
    # Process grades for all quarters
    sf9_updates = {}
    sf10_updates = {}
    for grade_info in student_grades:
        lrn, subject_idx, quarter, grade = grade_info
        if grade is None:
            continue
        if quarter <= 2:
            mapping = quarter_mappings[quarter]
            sf9_row = mapping['sf9_start_row'] + subject_idx
            if subject_idx in [3, 4, 5]:
                sf9_row += 1
            elif subject_idx in [6, 7]:
                sf9_row += 2
            elif subject_idx == 8:
                sf9_row += 2
            sf9_cell = f"{mapping['sf9_col']}{sf9_row}"
            sf10_cell = f"{mapping['sf10_col']}{mapping['sf10_start_row'] + subject_idx}"
            
            sf9_updates[sf9_cell] = grade
            sf10_updates[sf10_cell] = grade
        else:
            sf9_cell = direct_mappings.get((subject_idx, quarter))
            if sf9_cell:
                sf9_updates[sf9_cell] = grade
            mapping = quarter_mappings[quarter]
            sf10_cell = f"{mapping['sf10_col']}{mapping['sf10_start_row'] + subject_idx}"
            sf10_updates[sf10_cell] = grade
    for cell, value in sf9_updates.items():
        back_sf9.range(cell).value = value
    for cell, value in sf10_updates.items():
        front_sf10.range(cell).value = value
def main():
    start_time = time.time()
    sf9_folder, sf10_folder = create_folders()
    
    # Create and setup database
    conn = create_database()
    
    # Load MFQ file paths
    mfq_paths = ['MFQ1.xlsx', 'MFQ2.xlsx', 'MFQ3.xlsx', 'MFQ4.xlsx']
    
    # Extract data from Excel to SQLite
    print("Loading data from Excel to database...")
    load_data_from_excel_to_db(conn, mfq_paths)
    
    # Retrieve all student data from database - use pandas for efficiency
    print("Retrieving data from database...")
    students_df = pd.read_sql_query("SELECT * FROM students", conn)
    grades_df = pd.read_sql_query("SELECT * FROM grades", conn)
    
    student_data = list(students_df.itertuples(index=False, name=None))
    grades_data = list(grades_df.itertuples(index=False, name=None))
    
    # Process student files in parallel batches
    print(f"Processing {len(student_data)} students...")
    process_student_files(student_data, grades_data, sf9_folder, sf10_folder, 
                         max_workers=min(4, os.cpu_count()), batch_size=5)
    
    # Close database connection
    conn.close()
    
    end_time = time.time()
    print(f"Processing completed in {end_time - start_time:.2f} seconds.")

if __name__ == '__main__':
    main()