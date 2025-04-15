import os
import subprocess

# Get the current folder path
current_folder = os.path.dirname(os.path.abspath(__file__))

# Define the batch file name
bat_file = "opener.bat"  # Replace with your batch file name

# Construct the full path to the batch file
bat_file_path = os.path.join(current_folder, bat_file)

# Run the batch file
if os.path.exists(bat_file_path):
    subprocess.run(bat_file_path, shell=True)
else:
    print(f"Batch file '{bat_file}' not found in the folder.")
