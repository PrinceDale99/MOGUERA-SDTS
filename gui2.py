import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import os
import sys
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import subprocess
import time
from datetime import datetime
import win32com.client
import pythoncom

class ExcelPrinter:
    def __init__(self):
        self.excel = None
        
    def initialize_excel(self):
        pythoncom.CoInitialize()
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
    
    def print_excel_file(self, filepath):
        try:
            if not self.excel:
                self.initialize_excel()
                
            workbook = self.excel.Workbooks.Open(os.path.abspath(filepath))
            workbook.PrintOut()
            workbook.Close()
            return True
        except Exception as e:
            return str(e)
    
    def cleanup(self):
        if self.excel:
            self.excel.Quit()
            pythoncom.CoUninitialize()

class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
        
    def on_modified(self, event):
        if event.src_path.endswith('.xlsx'):
            self.callback()

class StudentManagementGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Student Management System")
        self.root.geometry("1000x800")
        
        # Configure grid
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        
        # Variables
        self.watchdog_active = False
        self.observer = None
        self.excel_printer = ExcelPrinter()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Create main frames
        left_frame = ctk.CTkFrame(self.root)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        right_frame = ctk.CTkFrame(self.root)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        # Left Frame - Student Data Management
        title_left = ctk.CTkLabel(left_frame, text="Student Data Management", font=("Arial", 20, "bold"))
        title_left.pack(pady=10)
        
        # Watchdog Toggle
        self.toggle_var = ctk.StringVar(value="off")
        toggle_switch = ctk.CTkSwitch(left_frame, text="Auto Update", 
                                    command=self.toggle_watchdog,
                                    variable=self.toggle_var, onvalue="on", offvalue="off")
        toggle_switch.pack(pady=10)
        
        # School Forms Button
        forms_btn = ctk.CTkButton(left_frame, text="Open School Forms", 
                                 command=self.show_school_forms,
                                 fg_color="#7289da", hover_color="#5b6eae")
        forms_btn.pack(pady=10)
        
        # Transfer Data Button
        transfer_btn = ctk.CTkButton(left_frame, text="Transfer data to Grade Master Form",
                                   command=lambda: self.run_script('run2.py'),
                                   fg_color="#7289da", hover_color="#5b6eae")
        transfer_btn.pack(pady=10)
        
        # Right Frame - Grade Encoding Assistant
        title_right = ctk.CTkLabel(right_frame, text="Grade Encoding Assistant", font=("Arial", 20, "bold"))
        title_right.pack(pady=10)
        
        # Edit Master Form Button
        edit_btn = ctk.CTkButton(right_frame, text="Edit Master Form",
                                command=self.show_quarter_selection,
                                fg_color="#7289da", hover_color="#5b6eae")
        edit_btn.pack(pady=10)
        
        # Finish & Encode Button
        finish_btn = ctk.CTkButton(right_frame, text="Finish & Encode",
                                 command=lambda: self.run_script('run3.py'),
                                 fg_color="#7289da", hover_color="#5b6eae")
        finish_btn.pack(pady=10)
        
        # Print Forms Button
        print_btn = ctk.CTkButton(right_frame, text="PRINT FORMS",
                                command=self.show_print_selection,
                                fg_color="#7289da", hover_color="#5b6eae")
        print_btn.pack(pady=10)
        
        # Console Output
        self.console = ctk.CTkTextbox(self.root, height=200, state="disabled")
        self.console.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="nsew")
        
        # Exit Button
        exit_btn = ctk.CTkButton(self.root, text="Exit Program",
                               command=self.exit_program,
                               fg_color="#7289da", hover_color="#5b6eae")
        exit_btn.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
    def log_message(self, message):
        self.console.configure(state="normal")
        self.console.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.console.configure(state="disabled")
        self.console.see("end")
        
    def toggle_watchdog(self):
        if self.toggle_var.get() == "on":
            self.start_watchdog()
        else:
            self.stop_watchdog()
    
    def start_watchdog(self):
        self.watchdog_active = True
        event_handler = ExcelFileHandler(self.on_excel_modified)
        self.observer = Observer()
        self.observer.schedule(event_handler, path='.', recursive=False)
        self.observer.start()
        self.log_message("Watchdog started - monitoring for Excel file changes")
    
    def stop_watchdog(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.watchdog_active = False
            self.log_message("Watchdog stopped")
    
    def on_excel_modified(self):
        self.log_message("Excel file modified - rerunning analysis")
        # Add your rerun logic here
        
    def show_school_forms(self):
        forms_window = ctk.CTkToplevel(self.root)
        forms_window.title("School Forms")
        forms_window.geometry("400x300")
        
        label = ctk.CTkLabel(forms_window, text="To edit the Master Form, open SF1")
        label.pack(pady=10)
        
        sf1_btn = ctk.CTkButton(forms_window, text="SF1",
                               command=lambda: os.startfile("sf1.xlsx"),
                               fg_color="#7289da", hover_color="#5b6eae")
        sf1_btn.pack(pady=5)
        
        sf5a_btn = ctk.CTkButton(forms_window, text="SF5A",
                                command=lambda: os.startfile("sf5a.xlsx"),
                                fg_color="#7289da", hover_color="#5b6eae")
        sf5a_btn.pack(pady=5)
        
        sf5b_btn = ctk.CTkButton(forms_window, text="SF5B",
                                command=lambda: os.startfile("sf5b.xlsx"),
                                fg_color="#7289da", hover_color="#5b6eae")
        sf5b_btn.pack(pady=5)
        
        close_btn = ctk.CTkButton(forms_window, text="Close",
                                 command=forms_window.destroy,
                                 fg_color="#7289da", hover_color="#5b6eae")
        close_btn.pack(pady=10)
        
    def show_quarter_selection(self):
        quarter_window = ctk.CTkToplevel(self.root)
        quarter_window.title("Quarter Selection")
        quarter_window.geometry("400x300")
        
        label = ctk.CTkLabel(quarter_window, text="Select Which Quarter to Edit")
        label.pack(pady=10)
        
        quarters = ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"]
        files = ["MFQ1.xlsx", "MFQ2.xlsx", "MFQ3.xlsx", "MFQ4.xlsx"]
        
        for q, f in zip(quarters, files):
            btn = ctk.CTkButton(quarter_window, text=q,
                              command=lambda file=f: os.startfile(file),
                              fg_color="#7289da", hover_color="#5b6eae")
            btn.pack(pady=5)
        
        close_btn = ctk.CTkButton(quarter_window, text="Close",
                                 command=quarter_window.destroy,
                                 fg_color="#7289da", hover_color="#5b6eae")
        close_btn.pack(pady=10)
    
    def print_file(self, filepath):
        if os.path.exists(filepath):
            result = self.excel_printer.print_excel_file(filepath)
            if result is True:
                self.log_message(f"Printing {filepath}")
            else:
                self.log_message(f"Error printing {filepath}: {result}")
        else:
            self.log_message(f"Error: File not found - {filepath}")
    
    def print_directory(self, directory):
        if os.path.exists(directory):
            for file in os.listdir(directory):
                if file.endswith('.xlsx'):
                    self.print_file(os.path.join(directory, file))
        else:
            self.log_message(f"Error: Directory not found - {directory}")
        
    def show_print_selection(self):
        print_window = ctk.CTkToplevel(self.root)
        print_window.title("Print Selection")
        print_window.geometry("400x400")
        
        options = {
            "School Form 1": "sf1.xlsx",
            "School Form 5a": "sf5a.xlsx",
            "School Form 5b": "sf5b.xlsx",
            "School Forms 9": "/SF9SF10/SF9/",
            "School Forms 10": "/SF9SF10/SF10/"
        }
        
        vars_dict = {}
        for option in options:
            var = tk.BooleanVar()
            vars_dict[option] = var
            checkbox = ctk.CTkCheckBox(print_window, text=option, variable=var)
            checkbox.pack(pady=5)
        
        def print_selected():
            selected = [opt for opt, var in vars_dict.items() if var.get()]
            if not selected:
                messagebox.showerror("Error", "Select at least 1 file!")
                return
            
            self.log_message("Starting print job...")
            for option in selected:
                path = options[option]
                if path.endswith('.xlsx'):
                    self.print_file(path)
                else:
                    self.print_directory(path)
            
            self.log_message(f"Print jobs completed for: {', '.join(selected)}")
        
        print_btn = ctk.CTkButton(print_window, text="Print",
                                command=print_selected,
                                fg_color="#7289da", hover_color="#5b6eae")
        print_btn.pack(pady=10)
        
        close_btn = ctk.CTkButton(print_window, text="Close",
                                 command=print_window.destroy,
                                 fg_color="#7289da", hover_color="#5b6eae")
        close_btn.pack(pady=10)
    
    def run_script(self, script_name):
        try:
            result = subprocess.run([sys.executable, script_name], 
                                  capture_output=True, text=True)
            self.log_message(f"Running {script_name}...")
            if result.stdout:
                self.log_message(result.stdout)
            if result.stderr:
                self.log_message(f"Error: {result.stderr}")
        except Exception as e:
            self.log_message(f"Error running {script_name}: {str(e)}")
    
    def exit_program(self):
        if self.watchdog_active:
            self.stop_watchdog()
        self.excel_printer.cleanup()
        self.root.quit()
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = StudentManagementGUI()
    app.run()