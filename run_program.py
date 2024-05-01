import os
import subprocess

def list_python_files():
    """
    Lists all the .py files in the current directory, excluding run_program.py.
    """
    current_file = os.path.basename(__file__)
    python_files = [f for f in os.listdir() if f.endswith('.py') and f != current_file]
    return python_files

def select_file(files):
    """
    Prompts the user to select a file from the list of files.
    """
    print("Select a Python file to run:")
    for i, file in enumerate(files):
        print(f"{i+1}. {file}")
    
    choice = input("Enter the number of the file you want to run: ")
    try:
        choice = int(choice)
        if 1 <= choice <= len(files):
            return files[choice-1]
        else:
            print("Invalid choice. Please enter a valid number.")
            return select_file(files)
    except ValueError:
        print("Invalid input. Please enter a number.")
        return select_file(files)

def run_selected_file(file_name):
    """
    Runs the selected Python file using subprocess.
    """
    try:
        subprocess.run(["python", file_name], check=True)
        if file_name == "PuretoResearchfish.py":
            subprocess.run(["python", "PureAddFunder.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
    except FileNotFoundError:
        print("Error: Python interpreter not found. Please ensure Python is installed.")

if __name__ == "__main__":
    files = list_python_files()
    if not files:
        print("No other .py files found in the current directory.")
    else:
        selected_file = select_file(files)
        run_selected_file(selected_file)
