import unittest
import os
import shutil
from src.file_saver import FileSaver

class TestFileSaver(unittest.TestCase):

    def setUp(self):
        # Create a temporary directory for testing
        self.test_base_directory = "tests/temp"
        os.makedirs(self.test_base_directory, exist_ok=True)
        self.file_saver = FileSaver(base_directory=self.test_base_directory)

    def tearDown(self):
        # Clean up the temporary directory after tests
        if os.path.exists(self.test_base_directory):
            shutil.rmtree(self.test_base_directory)

    def test_save_email(self):
        # Test saving an email in .eml format
        email = type('Email', (object,), {'subject': 'Test Email', 'plain_text_body': 'This is a test email.'})()
        folder_path = "emails"
        self.file_saver.save_email(email, folder_path, format='eml')

        # Verify the file was created
        expected_file_path = os.path.join(self.test_base_directory, folder_path, "Test Email.eml")
        self.assertTrue(os.path.exists(expected_file_path))

        # Verify the file content
        with open(expected_file_path, 'r') as file:
            content = file.read()
        self.assertIn("This is a test email.", content)

    def test_save_attachment(self):
        # Test saving an attachment
        attachment = type('Attachment', (object,), {'filename': 'test_attachment.txt', 'content': b'This is a test attachment.'})()
        folder_path = "attachments"
        self.file_saver.save_attachment(attachment, folder_path)

        # Verify the file was created
        expected_file_path = os.path.join(self.test_base_directory, folder_path, "test_attachment.txt")
        self.assertTrue(os.path.exists(expected_file_path))

        # Verify the file content
        with open(expected_file_path, 'rb') as file:
            content = file.read()
        self.assertEqual(content, b'This is a test attachment.')

    def test_save_email_with_structure(self):
        # Test saving an email while maintaining folder structure
        email = type('Email', (object,), {'subject': 'Structured Email', 'plain_text_body': 'This is a structured email.'})()
        folder_path = "structured/emails"
        self.file_saver.save_email(email, folder_path, format='eml')

        # Verify the file was created
        expected_file_path = os.path.join(self.test_base_directory, folder_path, "Structured Email.eml")
        self.assertTrue(os.path.exists(expected_file_path))

        # Verify the file content
        with open(expected_file_path, 'r') as file:
            content = file.read()
        self.assertIn("This is a structured email.", content)

if __name__ == '__main__':
    unittest.main()