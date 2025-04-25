import unittest
import os
from src.utils.helpers import (
    get_file_extension,
    create_directory,
    format_email_subject,
    save_file,
    load_file
)

class TestHelpers(unittest.TestCase):

    def setUp(self):
        # Create a temporary directory for testing
        self.test_directory = "tests/temp"
        os.makedirs(self.test_directory, exist_ok=True)

    def tearDown(self):
        # Clean up the temporary directory after tests
        if os.path.exists(self.test_directory):
            for root, dirs, files in os.walk(self.test_directory, topdown=False):
                for file in files:
                    os.remove(os.path.join(root, file))
                for dir in dirs:
                    os.rmdir(os.path.join(root, dir))
            os.rmdir(self.test_directory)

    def test_get_file_extension(self):
        # Test getting file extensions
        self.assertEqual(get_file_extension("file.txt"), "txt")
        self.assertEqual(get_file_extension("archive.tar.gz"), "gz")
        self.assertEqual(get_file_extension("no_extension"), "")
        self.assertEqual(get_file_extension(".hiddenfile"), "")

    def test_create_directory(self):
        # Test creating a directory
        test_path = os.path.join(self.test_directory, "new_folder")
        create_directory(test_path)
        self.assertTrue(os.path.exists(test_path))
        self.assertTrue(os.path.isdir(test_path))

    def test_format_email_subject(self):
        # Test formatting email subjects
        self.assertEqual(format_email_subject("Test/Email:Subject"), "Test_Email_Subject")
        self.assertEqual(format_email_subject("Invalid|Characters?"), "Invalid_Characters_")
        self.assertEqual(format_email_subject("NormalSubject"), "NormalSubject")

    def test_save_file_and_load_file(self):
        # Test saving and loading a file
        test_file_path = os.path.join(self.test_directory, "test_file.txt")
        test_content = b"This is a test file."

        # Save the file
        save_file(test_file_path, test_content)
        self.assertTrue(os.path.exists(test_file_path))

        # Load the file
        loaded_content = load_file(test_file_path)
        self.assertEqual(loaded_content, test_content)

if __name__ == '__main__':
    unittest.main()