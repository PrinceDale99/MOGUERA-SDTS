import subprocess
import sys
import os
import logging
from typing import List

def setup_logging():
    """Configure logging for the script runner."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

def run_script(script_path: str) -> bool:
    """
    Run a Python script and return whether it was successful.
    
    Args:
        script_path (str): Path to the Python script to run
        
    Returns:
        bool: True if script ran successfully, False otherwise
    """
    try:
        logging.info(f"Starting execution of {script_path}")
        
        # Get the Python executable path that's currently running
        python_executable = sys.executable
        
        # Run the script using the same Python interpreter
        result = subprocess.run(
            [python_executable, script_path],
            capture_output=True,
            text=True,
            check=True
        )
        
        # Log the output
        if result.stdout:
            logging.info(f"Output from {script_path}:\n{result.stdout}")
            
        logging.info(f"Successfully completed execution of {script_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running {script_path}:")
        logging.error(f"Exit code: {e.returncode}")
        if e.stdout:
            logging.error(f"Output:\n{e.stdout}")
        if e.stderr:
            logging.error(f"Error output:\n{e.stderr}")
        return False
        
    except Exception as e:
        logging.error(f"Unexpected error running {script_path}: {str(e)}")
        return False

def main():
    """Main function to run scripts sequentially."""
    setup_logging()
    
    # List of scripts to run in order
    scripts = ['schooldata.py', 'grade.py']
    
    # Verify all scripts exist
    for script in scripts:
        if not os.path.exists(script):
            logging.error(f"Script not found: {script}")
            sys.exit(1)
    
    # Run scripts sequentially
    for script in scripts:
        success = run_script(script)
        if not success:
            logging.error(f"Failed to run {script}. Stopping execution.")
            sys.exit(1)
    
    logging.info("All scripts completed successfully")

if __name__ == "__main__":
    main()