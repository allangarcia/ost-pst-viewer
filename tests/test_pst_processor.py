import os
import shutil
import tempfile
import unittest

from src.pst_processor import PSTProcessor


class TestPSTProcessor(unittest.TestCase):
    """
    Test cases for the PSTProcessor class.

    This test suite verifies the functionality of PST/OST file discovery
    and validation operations.
    """

    def setUp(self):
        """
        Set up test fixtures before each test method.

        Creates a temporary directory for testing and initializes a PSTProcessor
        instance.
        """
        # Create a temporary directory for testing
        self.test_temp_dir = tempfile.mkdtemp()
        self.test_pst_dir = os.path.join(self.test_temp_dir, "test_pst_files")
        self.processor = PSTProcessor(pst_directory=self.test_pst_dir)

    def tearDown(self):
        """
        Clean up test fixtures after each test method.

        Removes the temporary directory and all its contents.
        """
        # Clean up the temporary directory after tests
        if os.path.exists(self.test_temp_dir):
            shutil.rmtree(self.test_temp_dir)

    def test_find_pst_files_empty_directory(self):
        """
        Test finding PST files in an empty directory.

        Verifies that:
        - Empty list is returned when no PST/OST files exist
        - Directory is created if it doesn't exist
        """
        pst_files = self.processor.find_pst_files()
        self.assertEqual(pst_files, [])
        self.assertTrue(os.path.exists(self.test_pst_dir))

    def test_find_pst_files_with_files(self):
        """
        Test finding PST files when files exist.

        Verifies that:
        - PST and OST files are found correctly
        - Results are sorted alphabetically
        - Non-PST files are ignored
        """
        # Create test directory and files
        os.makedirs(self.test_pst_dir, exist_ok=True)

        # Create test files
        test_files = [
            "test1.pst",
            "test2.ost",
            "archive.pst",
            "readme.txt",  # Should be ignored
            "data.csv",  # Should be ignored
        ]

        for filename in test_files:
            filepath = os.path.join(self.test_pst_dir, filename)
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("test content")

        pst_files = self.processor.find_pst_files()

        # Should find only PST/OST files, sorted
        expected_files = [
            os.path.join(self.test_pst_dir, "archive.pst"),
            os.path.join(self.test_pst_dir, "test1.pst"),
            os.path.join(self.test_pst_dir, "test2.ost"),
        ]

        self.assertEqual(pst_files, expected_files)

    def test_validate_pst_file_valid(self):
        """
        Test validation of valid PST/OST files.

        Verifies that:
        - Valid PST files return True
        - Valid OST files return True
        """
        # Create test files
        os.makedirs(self.test_pst_dir, exist_ok=True)

        pst_file = os.path.join(self.test_pst_dir, "test.pst")
        ost_file = os.path.join(self.test_pst_dir, "test.ost")

        with open(pst_file, "w", encoding="utf-8") as f:
            f.write("test content")
        with open(ost_file, "w", encoding="utf-8") as f:
            f.write("test content")

        self.assertTrue(self.processor.validate_pst_file(pst_file))
        self.assertTrue(self.processor.validate_pst_file(ost_file))

    def test_validate_pst_file_invalid(self):
        """
        Test validation of invalid files.

        Verifies that:
        - Non-existent files return False
        - Files with wrong extensions return False
        - Empty/None paths return False
        """
        # Test non-existent file
        self.assertFalse(self.processor.validate_pst_file("/non/existent/file.pst"))

        # Test wrong extension
        os.makedirs(self.test_pst_dir, exist_ok=True)
        txt_file = os.path.join(self.test_pst_dir, "test.txt")
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write("test content")

        self.assertFalse(self.processor.validate_pst_file(txt_file))

        # Test empty/None paths
        self.assertFalse(self.processor.validate_pst_file(""))
        self.assertFalse(self.processor.validate_pst_file(None))

    def test_get_pst_files_count(self):
        """
        Test counting PST files in directory.

        Verifies that:
        - Correct count is returned
        - Count is 0 for empty directories
        """
        # Test empty directory
        self.assertEqual(self.processor.get_pst_files_count(), 0)

        # Test with files
        os.makedirs(self.test_pst_dir, exist_ok=True)

        # Create test PST files
        for i in range(3):
            filepath = os.path.join(self.test_pst_dir, f"test{i}.pst")
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("test content")

        self.assertEqual(self.processor.get_pst_files_count(), 3)

    def test_ensure_pst_directory_exists(self):
        """
        Test directory creation functionality.

        Verifies that:
        - Directory is created when it doesn't exist
        - Returns True for successful creation
        - Returns True if directory already exists
        """
        # Directory shouldn't exist initially
        self.assertFalse(os.path.exists(self.test_pst_dir))

        # Should create directory and return True
        result = self.processor.ensure_pst_directory_exists()
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.test_pst_dir))

        # Should return True if directory already exists
        result = self.processor.ensure_pst_directory_exists()
        self.assertTrue(result)


if __name__ == "__main__":
    unittest.main()
