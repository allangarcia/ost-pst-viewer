import unittest
import os
import shutil
from src.folder_structure import FolderStructure

class TestFolderStructure(unittest.TestCase):

    def setUp(self):
        # Create a temporary base directory for testing
        self.test_base_directory = "tests/temp"
        os.makedirs(self.test_base_directory, exist_ok=True)
        self.folder_structure = FolderStructure(base_path=self.test_base_directory)

    def test_create_folder(self):
        # Test creating a single folder
        folder_name = "test_folder"
        result_path = self.folder_structure.create_folder(folder_name)

        # Verify the folder was created
        expected_path = os.path.join(self.test_base_directory, folder_name)
        self.assertTrue(os.path.exists(expected_path))
        self.assertEqual(result_path, expected_path)

    def test_maintain_structure(self):
        # Test maintaining folder structure
        original_structure = ["folder1", "folder2/subfolder1", "folder2/subfolder2"]
        self.folder_structure.maintain_structure(original_structure)

        # Verify all folders were created
        for folder in original_structure:
            expected_path = os.path.join(self.test_base_directory, folder)
            self.assertTrue(os.path.exists(expected_path))

    def tearDown(self):
        # Clean up the temporary base directory after tests
        if os.path.exists(self.test_base_directory):
            shutil.rmtree(self.test_base_directory)

if __name__ == '__main__':
    unittest.main()