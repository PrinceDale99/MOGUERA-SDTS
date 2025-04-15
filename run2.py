import subprocess
import sys
import os

def run_script(script_name):
    """
    Run a Python script and wait for it to complete.
    Returns True if successful, False if there was an error.
    """
    try:
        # Get the Python executable path that's currently running
        python_executable = sys.executable
        
        # Check if the script exists
        if not os.path.exists(script_name):
            print(f"Error: {script_name} not found in the current directory")
            return False
        
        # Run the script using the same Python interpreter
        process = subprocess.run(
            [python_executable, script_name],
            check=True,
            text=True,
            capture_output=True
        )
        
        # Print the script's output
        if process.stdout:
            print(f"Output from {script_name}:")
            print(process.stdout)
            
        if process.stderr:
            print(f"Errors/warnings from {script_name}:")
            print(process.stderr)
            
        print(f"Successfully completed {script_name}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_name}:")
        print(f"Exit code: {e.returncode}")
        print(f"Error output: {e.stderr}")
        return False
    except Exception as e:
        print(f"Unexpected error running {script_name}:")
        print(str(e))
        return False

def main():
    # List of scripts to run in order
    scripts = ['squish.py', 'brock.py']
    
    for script in scripts:
        print(f"\nStarting {script}...")
        success = run_script(script)
        
        if not success:
            print(f"\nExecution stopped due to error in {script}")
            break
        print(f"Completed {script}")

if __name__ == "__main__":
    main()