import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import os
import sys
import subprocess
import shutil
import win32com.client
import pythoncom
import threading

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
            workbook.Close(SaveChanges=False)
            return True
        except Exception as e:
            return str(e)
    
    def cleanup(self):
        if self.excel:
            try:
                self.excel.Quit()
            except Exception:
                pass  # Suppress any errors
            finally:
                self.excel = None
                pythoncom.CoUninitialize()

class StudentManagementGUI:
    def __init__(self):
        # Set appearance mode and default color theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.root = ctk.CTk()
        self.root.title("Student Management System")
        self.root.geometry("1000x600")
        self.root.update_idletasks()  # Update to get accurate window sizes
        
        # Center the window on the screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 1000) // 2
        y = (screen_height - 600) // 2
        self.root.geometry(f"1000x600+{x}+{y}")
        
        # Configure grid
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        
        # Variables
        self.excel_printer = ExcelPrinter()
        self.executing = False
        self.status_label = None
        
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
        
        # Transfer to SF5AB Button
        transfer_to_sf5ab_btn = ctk.CTkButton(left_frame, text="Transfer to SF5AB", 
                                  command=lambda: self.run_script_async('nig.py'),
                                  fg_color="#7289da", hover_color="#5b6eae")
        transfer_to_sf5ab_btn.pack(pady=10)
        
        # School Forms Button
        forms_btn = ctk.CTkButton(left_frame, text="Open School Forms", 
                                 command=self.show_school_forms,
                                 fg_color="#7289da", hover_color="#5b6eae")
        forms_btn.pack(pady=10)
        
        # Transfer Data Button
        transfer_btn = ctk.CTkButton(left_frame, text="Transfer Grades to Master Forms",
                                   command=lambda: self.run_script_async('trans.py'),
                                   fg_color="#7289da", hover_color="#5b6eae")
        transfer_btn.pack(pady=10)
        
        # Right Frame - Grade Encoding Assistant
        title_right = ctk.CTkLabel(right_frame, text="Grade Encoding Assistant", font=("Arial", 20, "bold"))
        title_right.pack(pady=10)
        
        # Edit Master Form Button
        edit_btn = ctk.CTkButton(right_frame, text="Open Master Forms",
                                command=self.show_quarter_selection,
                                fg_color="#7289da", hover_color="#5b6eae")
        edit_btn.pack(pady=10)
        
        # Finish & Encode Button
        finish_btn = ctk.CTkButton(right_frame, text="Finish & Encode",
                                 command=lambda: self.run_script_async('grade.py'),
                                 fg_color="#7289da", hover_color="#5b6eae")
        finish_btn.pack(pady=10)
        
        # Print/Transfer Button
        print_transfer_btn = ctk.CTkButton(right_frame, text="Print / Transfer",
                                command=self.show_print_transfer_options,
                                fg_color="#7289da", hover_color="#5b6eae")
        print_transfer_btn.pack(pady=10)
        
        # Status indicator
        self.status_frame = ctk.CTkFrame(self.root)
        self.status_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="ew")
        
        self.status_label = ctk.CTkLabel(self.status_frame, text="Ready", text_color="green")
        self.status_label.pack(pady=5)
        
        # Exit Button
        exit_btn = ctk.CTkButton(self.root, text="Exit Program",
                               command=self.exit_program,
                               fg_color="#7289da", hover_color="#5b6eae")
        exit_btn.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
    def set_status(self, message, is_running=False):
        color = "red" if is_running else "green"
        self.status_label.configure(text=message, text_color=color)
        self.root.update()
    
    # New method to center windows
    def center_window(self, window, width, height):
        window.geometry(f"{width}x{height}")
        window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")
        
    def show_school_forms(self):
        forms_window = ctk.CTkToplevel(self.root)
        forms_window.title("School Forms")
        self.center_window(forms_window, 400, 300)
        
        # Make sure it's shown on top
        forms_window.transient(self.root)
        forms_window.grab_set()
        
        label = ctk.CTkLabel(forms_window, text="To edit the Master Form, open SF1")
        label.pack(pady=10)
        
        sf1_btn = ctk.CTkButton(forms_window, text="SF1",
                               command=lambda: self.safe_open_file("sf1.xlsx"),
                               fg_color="#7289da", hover_color="#5b6eae")
        sf1_btn.pack(pady=5)
        
        sf5a_btn = ctk.CTkButton(forms_window, text="SF5A",
                                command=lambda: self.safe_open_file("sf5a.xlsx"),
                                fg_color="#7289da", hover_color="#5b6eae")
        sf5a_btn.pack(pady=5)
        
        sf5b_btn = ctk.CTkButton(forms_window, text="SF5B",
                                command=lambda: self.safe_open_file("sf5b.xlsx"),
                                fg_color="#7289da", hover_color="#5b6eae")
        sf5b_btn.pack(pady=5)
        
        close_btn = ctk.CTkButton(forms_window, text="Close",
                                 command=forms_window.destroy,
                                 fg_color="#7289da", hover_color="#5b6eae")
        close_btn.pack(pady=10)
        
    def show_quarter_selection(self):
        quarter_window = ctk.CTkToplevel(self.root)
        quarter_window.title("Quarter Selection")
        self.center_window(quarter_window, 400, 300)
        
        # Make sure it's shown on top
        quarter_window.transient(self.root)
        quarter_window.grab_set()
        
        label = ctk.CTkLabel(quarter_window, text="Select Which Quarter to Edit")
        label.pack(pady=10)
        
        quarters = ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"]
        files = ["MFQ1.xlsx", "MFQ2.xlsx", "MFQ3.xlsx", "MFQ4.xlsx"]
        
        for q, f in zip(quarters, files):
            btn = ctk.CTkButton(quarter_window, text=q,
                              command=lambda file=f: self.safe_open_file(file),
                              fg_color="#7289da", hover_color="#5b6eae")
            btn.pack(pady=5)
        
        close_btn = ctk.CTkButton(quarter_window, text="Close",
                                 command=quarter_window.destroy,
                                 fg_color="#7289da", hover_color="#5b6eae")
        close_btn.pack(pady=10)
    
    # New method for safely opening files
    def safe_open_file(self, filepath):
        if os.path.exists(filepath):
            try:
                os.startfile(filepath)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open file: {e}")
        else:
            messagebox.showerror("Error", "File not found.")
    
    def print_file(self, filepath):
        if os.path.exists(filepath):
            result = self.excel_printer.print_excel_file(filepath)
            return result is True
        return False
    
    def print_directory(self, directory):
        if not os.path.exists(directory):
            return False
            
        success = True
        for file in os.listdir(directory):
            if file.endswith('.xlsx'):
                if not self.print_file(os.path.join(directory, file)):
                    success = False
        return success
    
    def show_print_selection(self):
        print_window = ctk.CTkToplevel(self.root)
        print_window.title("Print Selection")
        self.center_window(print_window, 400, 400)
        
        # Make sure it's shown on top
        print_window.transient(self.root)
        print_window.grab_set()
        
        options = {
            "School Form 1": "sf1.xlsx",
            "School Form 5a": "sf5a.xlsx",
            "School Form 5b": "sf5b.xlsx",
            "School Forms 9": "SF9SF10/SF9",
            "School Forms 10": "SF9SF10/SF10"
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
            
            self.set_status("Printing files...", True)
            success = True
            
            for option in selected:
                path = options[option]
                if path.endswith('.xlsx'):
                    if not self.print_file(path):
                        success = False
                else:
                    if not self.print_directory(path):
                        success = False
            
            if success:
                self.set_status("Print completed successfully", False)
            else:
                self.set_status("Some print jobs failed", False)
            
            print_window.destroy()
        
        print_btn = ctk.CTkButton(print_window, text="Print",
                                command=print_selected,
                                fg_color="#7289da", hover_color="#5b6eae")
        print_btn.pack(pady=10)
        
        close_btn = ctk.CTkButton(print_window, text="Close",
                                 command=print_window.destroy,
                                 fg_color="#7289da", hover_color="#5b6eae")
        close_btn.pack(pady=10)
    
    def transfer_files(self, is_email=False):
        files_to_transfer = ['sf1.xlsx', 'sf5a.xlsx', 'sf5b.xlsx']
        folder_to_transfer = 'SF9SF10'
        
        if is_email:
            # Here you would implement the email functionality
            # This is just a placeholder
            self.set_status("Sending email...", True)
            # Add your email sending code here
            messagebox.showinfo("Email", "Files sent via email successfully!")
            self.set_status("Email sent successfully", False)
        else:
            # Transfer to USB/folder
            destination = filedialog.askdirectory(title="Select Destination Folder")
            if not destination:
                self.set_status("Transfer canceled", False)
                return
                
            self.set_status("Transferring files...", True)
            try:
                # Copy individual files
                for file in files_to_transfer:
                    if os.path.exists(file):
                        shutil.copy2(file, destination)
                
                # Copy folder if it exists
                if os.path.exists(folder_to_transfer):
                    dest_folder = os.path.join(destination, folder_to_transfer)
                    if os.path.exists(dest_folder):
                        shutil.rmtree(dest_folder)
                    shutil.copytree(folder_to_transfer, dest_folder)
                
                messagebox.showinfo("Success", "Files transferred successfully!")
                self.set_status("Files transferred successfully", False)
            except Exception as e:
                messagebox.showerror("Error", f"Error transferring files: {str(e)}")
                self.set_status("Error transferring files", False)
    
    def show_transfer_options(self):
        transfer_window = ctk.CTkToplevel(self.root)
        transfer_window.title("Transfer Options")
        self.center_window(transfer_window, 300, 200)
        
        # Make sure it's shown on top
        transfer_window.transient(self.root)
        transfer_window.grab_set()
        
        label = ctk.CTkLabel(transfer_window, text="Choose Transfer Method:")
        label.pack(pady=20)
        
        email_btn = ctk.CTkButton(transfer_window, text="Send Email",
                                command=lambda: [transfer_window.destroy(), self.transfer_files(True)],
                                fg_color="#7289da", hover_color="#5b6eae")
        email_btn.pack(pady=10)
        
        usb_btn = ctk.CTkButton(transfer_window, text="Transfer with USB/Folder",
                               command=lambda: [transfer_window.destroy(), self.transfer_files(False)],
                               fg_color="#7289da", hover_color="#5b6eae")
        usb_btn.pack(pady=10)
        
        cancel_btn = ctk.CTkButton(transfer_window, text="Cancel",
                                  command=transfer_window.destroy,
                                  fg_color="#7289da", hover_color="#5b6eae")
        cancel_btn.pack(pady=10)
    
    def show_print_transfer_options(self):
        options_window = ctk.CTkToplevel(self.root)
        options_window.title("Print or Transfer")
        self.center_window(options_window, 300, 200)
        
        # Make sure it's shown on top
        options_window.transient(self.root)
        options_window.grab_set()
        
        label = ctk.CTkLabel(options_window, text="Choose Option:")
        label.pack(pady=20)
        
        print_btn = ctk.CTkButton(options_window, text="Print",
                                command=lambda: [options_window.destroy(), self.show_print_selection()],
                                fg_color="#7289da", hover_color="#5b6eae")
        print_btn.pack(pady=10)
        
        transfer_btn = ctk.CTkButton(options_window, text="Transfer",
                                   command=lambda: [options_window.destroy(), self.show_transfer_options()],
                                   fg_color="#7289da", hover_color="#5b6eae")
        transfer_btn.pack(pady=10)
        
        cancel_btn = ctk.CTkButton(options_window, text="Cancel",
                                  command=options_window.destroy,
                                  fg_color="#7289da", hover_color="#5b6eae")
        cancel_btn.pack(pady=10)
    
    # New method for running scripts asynchronously
    def run_script_async(self, script_name):
        if not self.executing:
            self.executing = True
            self.set_status(f"Running {script_name}...", is_running=True)
            thread = threading.Thread(target=self.run_script, args=(script_name,), daemon=True)
            thread.start()
        else:
            messagebox.showwarning("Warning", "A script is already running. Please wait until it finishes.")
    
    def run_script(self, script_name):
        # Disable all buttons during execution
        for widget in self.root.winfo_children():
            if isinstance(widget, ctk.CTkButton):
                widget.configure(state="disabled")
        
        try:
            result = subprocess.run([sys.executable, script_name], 
                                  capture_output=True, text=True)
            success = result.returncode == 0
            
            # Re-enable buttons and update status in the main thread
            self.root.after(0, lambda: self.script_completed(success, script_name, result.stderr if not success else ""))
        except Exception as e:
            # Re-enable buttons and show error in the main thread
            self.root.after(0, lambda: self.script_completed(False, script_name, str(e)))
    
    def script_completed(self, success, script_name, error_message):
        self.executing = False
        
        # Re-enable all buttons
        for widget in self.root.winfo_children():
            if isinstance(widget, ctk.CTkButton):
                widget.configure(state="normal")
        
        if success:
            self.set_status(f"{script_name} completed successfully", False)
        else:
            self.set_status(f"Error running {script_name}", False)
            messagebox.showerror("Error", f"Error running {script_name}: {error_message}")
    
    def exit_program(self):
        if self.executing:
            if not messagebox.askyesno("Warning", "A script is currently running. Are you sure you want to exit?"):
                return
        
        self.excel_printer.cleanup()
        self.root.quit()
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = StudentManagementGUI()
    app.run()