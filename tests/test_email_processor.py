import unittest
from src.email_processor import EmailProcessor

class TestEmailProcessor(unittest.TestCase):

    def setUp(self):
        # Provide a valid file path for testing
        self.sample_pst_file = 'tests/sample.pst'  # Replace with the actual path to a sample PST file
        self.processor = EmailProcessor(file_path=self.sample_pst_file)

    def test_load_emails(self):
        # Test loading emails from a sample OST/PST file
        emails = self.processor.load_emails()
        self.assertIsInstance(emails, list)
        self.assertGreater(len(emails), 0)

    def test_save_as_eml(self):
        # Test saving an email as .eml format
        email = type('Email', (object,), {'subject': 'Test Email', 'plain_text_body': 'This is a test.'})()
        output_path = 'tests/output'
        self.processor.save_as_eml(email, output_path)
        self.assertTrue(True)  # Add assertions to verify the file was saved correctly

    def test_save_as_pdf(self):
        # Test saving an email as .pdf format
        email = type('Email', (object,), {'subject': 'Test Email', 'html_body': '<p>This is a test.</p>'})()
        output_path = 'tests/output'
        self.processor.save_as_pdf(email, output_path)
        self.assertTrue(True)  # Add assertions to verify the file was saved correctly

if __name__ == '__main__':
    unittest.main()