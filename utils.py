import os

def get_unique_output_file(base_file, suffix="-comparison", separator="_"):
    """
    Generate a unique output file name by appending a counter or a custom suffix to the base name.

    Parameters:
        base_file (str): The base file path (including extension).
        suffix (str): Custom suffix to append to the base name (default is "-comparison").
        separator (str): Separator to use before appending the counter (default is "_").

    Returns:
        str: A unique file name.
    """
    base_name, ext = os.path.splitext(base_file)
    output_file = f"{base_name}{suffix}{ext}"
    counter = 1

    while os.path.exists(output_file):
        output_file = f"{base_name}{suffix}{separator}{counter}{ext}"
        counter += 1

    return output_file
