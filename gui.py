import tkinter as tk
import customtkinter as ctk
import threading
import queue
import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import logging
from datetime import datetime
import os
import nig  # importing your existing script

class OutputRedirector:
    def __init__(self, queue):
        self.queue = queue

    def write(self, text):
        self.queue.put(text)

    def flush(self):
        pass

class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
        self.last_modified = time.time()
        self.cooldown = 1  # 1 second cooldown between processing

    def on_modified(self, event):
        if event.src_path.endswith('.xlsx'):
            current_time = time.time()
            if current_time - self.last_modified > self.cooldown:
                self.last_modified = current_time
                self.callback()

class MonitoringApp:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("Excel Monitor")
        self.app.geometry("800x600")
        ctk.set_appearance_mode("dark")
        
        # Configure grid
        self.app.grid_columnconfigure(0, weight=1)
        self.app.grid_rowconfigure(1, weight=1)
        
        # Create top frame for controls
        self.top_frame = ctk.CTkFrame(self.app)
        self.top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        # Toggle switch
        self.switch_var = ctk.StringVar(value="off")
        self.toggle_switch = ctk.CTkSwitch(
            self.top_frame,
            text="Monitor Excel Files",
            command=self.toggle_monitoring,
            variable=self.switch_var,
            onvalue="on",
            offvalue="off"
        )
        self.toggle_switch.pack(pady=10)
        
        # Terminal-like output
        self.output_frame = ctk.CTkFrame(self.app)
        self.output_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.output_frame.grid_columnconfigure(0, weight=1)
        self.output_frame.grid_rowconfigure(0, weight=1)
        
        self.output_text = ctk.CTkTextbox(
            self.output_frame,
            wrap=tk.WORD,
            font=("Courier", 12)
        )
        self.output_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Initialize variables
        self.observer = None
        self.running = False
        self.output_queue = queue.Queue()
        
        # Redirect stdout to our queue
        sys.stdout = OutputRedirector(self.output_queue)
        logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))
        
        # Start queue checking
        self.check_queue()

    def toggle_monitoring(self):
        if self.switch_var.get() == "on":
            self.start_monitoring()
        else:
            self.stop_monitoring()

    def start_monitoring(self):
        if not self.running:
            self.running = True
            self.log_message("Starting Excel file monitoring...")
            
            # Create observer for file system events
            self.observer = Observer()
            handler = ExcelFileHandler(self.process_excel_files)
            self.observer.schedule(handler, path=".", recursive=False)
            self.observer.start()

    def stop_monitoring(self):
        if self.running:
            self.running = False
            if self.observer:
                self.observer.stop()
                self.observer.join()
            self.log_message("Stopped Excel file monitoring.")

    def process_excel_files(self):
        try:
            self.log_message("Changes detected, processing Excel files...")
            nig.main()  # Run your existing script
            self.log_message("Processing completed successfully.")
        except Exception as e:
            self.log_message(f"Error processing files: {str(e)}")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.output_queue.put(f"[{timestamp}] {message}\n")

    def check_queue(self):
        while True:
            try:
                message = self.output_queue.get_nowait()
                self.output_text.insert("end", message)
                self.output_text.see("end")
            except queue.Empty:
                break
        self.app.after(100, self.check_queue)

    def run(self):
        self.app.mainloop()

if __name__ == "__main__":
    app = MonitoringApp()
    app.run()