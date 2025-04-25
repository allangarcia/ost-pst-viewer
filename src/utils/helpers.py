import os
import re

def get_file_extension(file_path):
    """Returns the file extension for a given file path."""
    _, ext = os.path.splitext(file_path)
    return ext[1:] if ext else ''

def create_directory(path):
    """Creates a directory if it does not exist."""
    if not os.path.exists(path):
        os.makedirs(path)

def format_email_subject(subject):
    """Formats the email subject to be filesystem-friendly."""
    return re.sub(r'[<>:"/\\|?*]', '_', subject)

def save_file(file_path, content):
    """Saves content to a specified file path."""
    with open(file_path, 'wb') as file:
        file.write(content)

def load_file(file_path):
    """Loads content from a specified file path."""
    with open(file_path, 'rb') as file:
        return file.read()