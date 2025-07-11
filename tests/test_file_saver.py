import os
import shutil
import unittest

from src.file_saver import FileSaver


class TestFileSaver(unittest.TestCase):
    """
    Test cases for the FileSaver class.

    This test suite verifies the functionality of file saving operations
    including email saving in different formats and attachment handling.
    Uses a temporary directory structure for isolated testing.
    """

    def setUp(self):
        """
        Set up test fixtures before each test method.

        Creates a temporary directory for testing and initializes a FileSaver
        instance.
        """
        # Create a temporary directory for testing
        self.test_base_directory = "tests/temp"
        os.makedirs(self.test_base_directory, exist_ok=True)
        self.file_saver = FileSaver(base_directory=self.test_base_directory)

    def tearDown(self):
        """
        Clean up test fixtures after each test method.

        Removes the temporary directory and all its contents.
        """
        # Clean up the temporary directory after tests
        if os.path.exists(self.test_base_directory):
            shutil.rmtree(self.test_base_directory)

    def test_save_email(self):
        """
        Test saving an email in .eml format.

        Verifies that:
        - The email file is created in the correct location
        - The file contains the expected content
        """
        # Test saving an email in .eml format
        email = type(
            "Email",
            (object,),
            {"subject": "Test Email", "plain_text_body": "This is a test email."},
        )()
        folder_path = "emails"
        self.file_saver.save_email(email, folder_path, format="eml")

        # Verify the file was created
        expected_file_path = os.path.join(
            self.test_base_directory, folder_path, "Test Email.eml"
        )
        self.assertTrue(os.path.exists(expected_file_path))

        # Verify the file content
        with open(expected_file_path, "r") as file:
            content = file.read()
        self.assertIn("This is a test email.", content)

    def test_save_attachment(self):
        """
        Test saving an attachment.

        Verifies that:
        - The attachment file is created with the correct name
        - The file contains the expected binary content
        """
        # Test saving an attachment
        attachment = type(
            "Attachment",
            (object,),
            {
                "filename": "test_attachment.txt",
                "content": b"This is a test attachment.",
            },
        )()
        folder_path = "attachments"
        self.file_saver.save_attachment(attachment, folder_path)

        # Verify the file was created
        expected_file_path = os.path.join(
            self.test_base_directory, folder_path, "test_attachment.txt"
        )
        self.assertTrue(os.path.exists(expected_file_path))

        # Verify the file content
        with open(expected_file_path, "rb") as file:
            content = file.read()
        self.assertEqual(content, b"This is a test attachment.")

    def test_save_email_with_structure(self):
        """
        Test saving an email while maintaining folder structure.

        Verifies that:
        - Nested folder structures are created properly
        - The email file is saved in the correct nested location
        - The file contains the expected content
        """
        # Test saving an email while maintaining folder structure
        email = type(
            "Email",
            (object,),
            {
                "subject": "Structured Email",
                "plain_text_body": "This is a structured email.",
            },
        )()
        folder_path = "structured/emails"
        self.file_saver.save_email(email, folder_path, format="eml")

        # Verify the file was created
        expected_file_path = os.path.join(
            self.test_base_directory, folder_path, "Structured Email.eml"
        )
        self.assertTrue(os.path.exists(expected_file_path))

        # Verify the file content
        with open(expected_file_path, "r") as file:
            content = file.read()
        self.assertIn("This is a structured email.", content)


if __name__ == "__main__":
    unittest.main()
