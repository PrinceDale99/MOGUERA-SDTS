import os
import subprocess
import sys

# Define the folder where offline packages are stored
PACKAGE_DIR = "offline_packages"
REQUIREMENTS_FILE = "requirements.txt"

def install_packages():
    if not os.path.exists(PACKAGE_DIR):
        print(f"Error: The folder '{PACKAGE_DIR}' does not exist. Please transfer the package files first.")
        return

    # Construct the pip install command
    cmd = [
        sys.executable, "-m", "pip", "install",
        "--no-index",
        "--find-links=" + PACKAGE_DIR,
        "-r", REQUIREMENTS_FILE
    ]

    # Run the installation command
    try:
        subprocess.run(cmd, check=True)
        print("\nAll packages have been installed successfully!")
    except subprocess.CalledProcessError:
        print("\nError: Failed to install some packages. Check the package versions and dependencies.")

if __name__ == "__main__":
    install_packages()
