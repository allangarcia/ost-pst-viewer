import unittest
import os
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock

# Add src directory to path for imports
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from email_processor import EmailProcessor


class TestEmailProcessorProcessEmails(unittest.TestCase):
    """Test cases for the new process_emails method in EmailProcessor."""

    def setUp(self):
        """Set up test fixtures before each test method."""
        self.test_dir = tempfile.mkdtemp()
        self.processor = EmailProcessor("test.pst")

    def tearDown(self):
        """Clean up after each test method."""
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)

    @patch('email_processor.PSTProcessor')
    def test_process_emails_invalid_file(self, mock_pst_processor_class):
        """Test process_emails with invalid PST file."""
        # Setup mock
        mock_pst_processor = Mock()
        mock_pst_processor.validate_pst_file.return_value = False
        mock_pst_processor_class.return_value = mock_pst_processor
        
        # Test
        result = self.processor.process_emails(
            output_dir=self.test_dir,
            output_format="eml"
        )
        
        # Verify
        self.assertFalse(result)
        mock_pst_processor.validate_pst_file.assert_called_once_with("test.pst")

    @patch('email_processor.PSTProcessor')
    @patch('email_processor.FileSaver')
    def test_process_emails_dry_run(self, mock_file_saver_class, mock_pst_processor_class):
        """Test process_emails in dry run mode."""
        # Setup mocks
        mock_pst_processor = Mock()
        mock_pst_processor.validate_pst_file.return_value = True
        mock_pst_processor_class.return_value = mock_pst_processor
        
        # Mock email data
        mock_email = Mock()
        mock_email.subject = "Test Email"
        email_data = {"message": mock_email, "folder_path": "Inbox"}
        
        # Mock load_emails to return test data
        with patch.object(self.processor, 'load_emails', return_value=[email_data]):
            with patch.object(self.processor, '_format_delivery_time', return_value="[2025-07-11]"):
                result = self.processor.process_emails(
                    output_dir=self.test_dir,
                    output_format="eml",
                    dry_run=True
                )
        
        # Verify
        self.assertTrue(result)
        # In dry run mode, FileSaver should not be instantiated
        mock_file_saver_class.assert_not_called()

    @patch('email_processor.PSTProcessor')
    @patch('email_processor.FileSaver')
    def test_process_emails_success_eml(self, mock_file_saver_class, mock_pst_processor_class):
        """Test successful email processing with EML format."""
        # Setup mocks
        mock_pst_processor = Mock()
        mock_pst_processor.validate_pst_file.return_value = True
        mock_pst_processor_class.return_value = mock_pst_processor
        
        mock_file_saver = Mock()
        mock_file_saver._create_full_path.return_value = "/test/path"
        mock_file_saver_class.return_value = mock_file_saver
        
        # Mock email data
        mock_email = Mock()
        mock_email.subject = "Test Email"
        mock_email.number_of_attachments = 0
        email_data = {"message": mock_email, "folder_path": "Inbox"}
        
        # Mock EmailProcessor methods
        with patch.object(self.processor, 'load_emails', return_value=[email_data]):
            with patch.object(self.processor, 'save_as_eml') as mock_save_eml:
                result = self.processor.process_emails(
                    output_dir=self.test_dir,
                    output_format="eml"
                )
        
        # Verify
        self.assertTrue(result)
        mock_save_eml.assert_called_once_with(mock_email, "/test/path")

    @patch('email_processor.PSTProcessor')
    @patch('email_processor.FileSaver')
    def test_process_emails_success_both_formats(self, mock_file_saver_class, mock_pst_processor_class):
        """Test successful email processing with both EML and PDF formats."""
        # Setup mocks
        mock_pst_processor = Mock()
        mock_pst_processor.validate_pst_file.return_value = True
        mock_pst_processor_class.return_value = mock_pst_processor
        
        mock_file_saver = Mock()
        mock_file_saver._create_full_path.return_value = "/test/path"
        mock_file_saver_class.return_value = mock_file_saver
        
        # Mock email data
        mock_email = Mock()
        mock_email.subject = "Test Email"
        mock_email.number_of_attachments = 0
        email_data = {"message": mock_email, "folder_path": "Inbox"}
        
        # Mock EmailProcessor methods
        with patch.object(self.processor, 'load_emails', return_value=[email_data]):
            with patch.object(self.processor, 'save_as_eml') as mock_save_eml:
                with patch.object(self.processor, 'save_as_pdf') as mock_save_pdf:
                    result = self.processor.process_emails(
                        output_dir=self.test_dir,
                        output_format="both"
                    )
        
        # Verify
        self.assertTrue(result)
        mock_save_eml.assert_called_once_with(mock_email, "/test/path")
        mock_save_pdf.assert_called_once_with(mock_email, "/test/path")

    @patch('email_processor.PSTProcessor')
    @patch('email_processor.FileSaver')
    def test_process_emails_with_attachments(self, mock_file_saver_class, mock_pst_processor_class):
        """Test email processing with attachments."""
        # Setup mocks
        mock_pst_processor = Mock()
        mock_pst_processor.validate_pst_file.return_value = True
        mock_pst_processor_class.return_value = mock_pst_processor
        
        mock_file_saver = Mock()
        mock_file_saver._create_full_path.return_value = "/test/path"
        mock_file_saver_class.return_value = mock_file_saver
        
        # Mock email with attachments
        mock_attachment = Mock()
        mock_attachment.name = "test.pdf"
        
        mock_email = Mock()
        mock_email.subject = "Test Email"
        mock_email.number_of_attachments = 1
        mock_email.get_attachment.return_value = mock_attachment
        email_data = {"message": mock_email, "folder_path": "Inbox"}
        
        # Mock os.makedirs and os.path.join
        with patch('email_processor.os.makedirs'):
            with patch('email_processor.os.path.join', return_value="/test/path/attachments"):
                with patch.object(self.processor, 'load_emails', return_value=[email_data]):
                    with patch.object(self.processor, 'save_as_eml'):
                        result = self.processor.process_emails(
                            output_dir=self.test_dir,
                            output_format="eml"
                        )
        
        # Verify
        self.assertTrue(result)
        mock_file_saver.save_attachment.assert_called_once_with(mock_attachment, "/test/path/attachments")

    def test_process_emails_method_signature(self):
        """Test that process_emails has the correct method signature."""
        import inspect
        sig = inspect.signature(self.processor.process_emails)
        params = list(sig.parameters.keys())
        
        expected_params = ['output_dir', 'output_format', 'verbose', 'dry_run']
        for param in expected_params:
            self.assertIn(param, params)
        
        # Check default values
        self.assertEqual(sig.parameters['output_format'].default, "eml")
        self.assertEqual(sig.parameters['verbose'].default, False)
        self.assertEqual(sig.parameters['dry_run'].default, False)


if __name__ == '__main__':
    unittest.main()
