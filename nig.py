import xlwings as xw
import logging
from typing import List, Dict, Tuple
import os

class SchoolFormProcessor:
    def __init__(self, sf1_path: str, sf5a_path: str, sf5b_path: str):
        """Initialize the processor with paths to the Excel files."""
        self.sf1_path = os.path.abspath(sf1_path)
        self.sf5a_path = os.path.abspath(sf5a_path)
        self.sf5b_path = os.path.abspath(sf5b_path)
        
        
        logging.basicConfig(level=logging.INFO,
                          format='%(asctime)s - %(levelname)s - %(message)s')
        
        
        self.SF1_RANGES = {
            'male': (11, 50),
            'female': (52, 91)
        }
        
        self.SF5A_RANGES = {
            'male': (13, 43),
            'female': (45, 66)
        }
        
        self.SF5B_RANGES = {
            'male': (15, 43),
            'female': (45, 65)
        }

    def read_sf1_data(self) -> Dict[str, List[Dict[str, str]]]:
        """
        Read data from SF1 and organize it by gender.
        Returns a dictionary with 'male' and 'female' lists of student records.
        """
        try:
            app = xw.App(visible=False)
            wb = app.books.open(self.sf1_path)
            sheet = wb.sheets[0]
            data = {'male': [], 'female': []}
            
            
            for gender, (start_row, end_row) in self.SF1_RANGES.items():
                for row in range(start_row, end_row + 1):
                    
                    lrn = sheet.range(f'B{row}').value
                    name = sheet.range(f'C{row}').value
                    sex = sheet.range(f'G{row}').value
                    
                    
                    if not lrn:
                        continue
                    
                    student = {
                        'lrn': str(lrn).strip() if lrn else '',
                        'name': str(name).strip() if name else '',
                        'gender': str(sex).strip() if sex else ''
                    }
                    
                    
                    if student['lrn'] and student['name']:
                        data[gender].append(student)
            
            wb.close()
            app.quit()
            return data
            
        except Exception as e:
            logging.error(f"Error reading SF1: {str(e)}")
            raise

    def write_to_sf5a(self, data: Dict[str, List[Dict[str, str]]], app):
        """Write organized data to SF5A."""
        try:
            wb = app.books.open(self.sf5a_path)
            sheet = wb.sheets[0]
            
            
            self._write_section(sheet, data['male'], 
                              self.SF5A_RANGES['male'],
                              'C', 'D')  
            
            
            self._write_section(sheet, data['female'],
                              self.SF5A_RANGES['female'],
                              'C', 'D')  
            
            wb.save()
            wb.close()
            logging.info("Successfully wrote data to SF5A")
            
        except Exception as e:
            logging.error(f"Error writing to SF5A: {str(e)}")
            raise

    def write_to_sf5b(self, data: Dict[str, List[Dict[str, str]]], app):
        """Write organized data to SF5B."""
        try:
            wb = app.books.open(self.sf5b_path)
            sheet = wb.sheets[0]
            
            
            self._write_section(sheet, data['male'],
                              self.SF5B_RANGES['male'],
                              'B', 'C')  
            
            
            self._write_section(sheet, data['female'],
                              self.SF5B_RANGES['female'],
                              'B', 'C')  
            
            wb.save()
            wb.close()
            logging.info("Successfully wrote data to SF5B")
            
        except Exception as e:
            logging.error(f"Error writing to SF5B: {str(e)}")
            raise

    def _write_section(self, sheet, data: List[Dict[str, str]], 
                      row_range: Tuple[int, int], 
                      lrn_col: str, name_start_col: str):
        """Helper method to write a section of data to a sheet."""
        start_row, end_row = row_range
        available_rows = end_row - start_row + 1
        
        
        if len(data) > available_rows:
            logging.warning(
                f"Data contains {len(data)} records but only {available_rows} rows available. "
                f"Some records will be truncated."
            )
        
        
        for idx, student in enumerate(data[:available_rows]):
            current_row = start_row + idx
            
            
            sheet.range(f'{lrn_col}{current_row}').value = student['lrn']
            
            
            sheet.range(f'{name_start_col}{current_row}').value = student['name']

    def process_all(self):
        """Process all forms in one go."""
        try:
            logging.info("Starting school form processing...")
            
            
            data = self.read_sf1_data()
            logging.info(f"Read {len(data['male'])} male and {len(data['female'])} female records from SF1")
            
            
            app = xw.App(visible=False)
            
            
            self.write_to_sf5a(data, app)
            self.write_to_sf5b(data, app)
            
            
            app.quit()
            
            logging.info("School form processing completed successfully")
            
        except Exception as e:
            logging.error(f"Error during processing: {str(e)}")
            raise

def main():
    
    required_files = ['sf1.xlsx', 'sf5a.xlsx', 'sf5b.xlsx']
    for file in required_files:
        if not os.path.exists(file):
            raise FileNotFoundError(f"Required file {file} not found in current directory")
    
    
    processor = SchoolFormProcessor('sf1.xlsx', 'sf5a.xlsx', 'sf5b.xlsx')
    processor.process_all()

if __name__ == "__main__":
    main()