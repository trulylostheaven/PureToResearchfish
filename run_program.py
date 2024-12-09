import os
import subprocess
from pathlib import Path


def list_python_files():
    """
    Lists all the .py files in the current directory, excluding the script itself and specific files.
    """
    current_file = Path(__file__).name
    excluded_files = {current_file, 'utils.py'}
    return [f for f in os.listdir() 
            if f.endswith('.py') and f not in excluded_files]


def select_file(files):
    """
    Prompts the user to select a file from the list of files.
    """
    print("Select a Python file to run:")
    for i, file in enumerate(files, start=1):
        print(f"{i}. {file}")
    
    while True:
        choice = input("Enter the number of the file you want to run: ")
        if choice.isdigit():
            choice = int(choice)
            if 1 <= choice <= len(files):
                return files[choice-1]
            else:
                print("Invalid choice. Please enter a number between 1 and", len(files))
        else:
            print("Invalid input. Please enter a valid number.")


def run_selected_file(file_name):
    """
    Runs the selected Python file using subprocess.
    """
    try:
        subprocess.run(["python", file_name], check=True)
        if file_name == "PuretoResearchfish.py":
            subprocess.run(["python", "PureAddFunder.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error while running {file_name}: {e}")
    except FileNotFoundError:
        print("Error: Python interpreter not found. Please ensure Python is installed.")
    except Exception as e:
        print(f"Unexpected error: {e}")


if __name__ == "__main__":
    files = list_python_files()
    if not files:
        print("No other .py files found in the current directory.")
    else:
        selected_file = select_file(files)
        run_selected_file(selected_file)
