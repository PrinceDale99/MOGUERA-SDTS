import os
import re
import subprocess

# Set your folder path
folder_path = r"C:\Users\Grace Martin\Documents\python\Gradetra\v8 optimization\Grade Transfer System (Consumer Edition)"

# Pattern to find import statements
import_pattern = re.compile(r"^\s*(?:import|from)\s+([a-zA-Z0-9_\.]+)")

# Set to store libraries
libraries = set()

# Scan all .py files
for file in os.listdir(folder_path):
    if file.endswith(".py"):
        with open(os.path.join(folder_path, file), "r", encoding="utf-8") as f:
            for line in f:
                match = import_pattern.match(line)
                if match:
                    lib = match.group(1).split('.')[0]  # Get the base package
                    libraries.add(lib)

# Function to get library version using pip
def get_library_version(library_name):
    try:
        result = subprocess.run(['pip', 'show', library_name], capture_output=True, text=True)
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if line.startswith("Version:"):
                    return line.split(":", 1)[1].strip()
        return "Not Installed"
    except Exception as e:
        return "Error: " + str(e)

# Save to a requirements.txt file with versions
with open("requirements_with_versions.txt", "w") as req_file:
    for lib in sorted(libraries):
        version = get_library_version(lib)
        req_file.write(f"{lib}=={version}\n")

print("Libraries with versions extracted and saved to requirements_with_versions.txt")