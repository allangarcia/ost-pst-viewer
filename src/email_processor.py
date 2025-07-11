import email as py_email
import os
import re
from datetime import datetime
from email import policy
from pathlib import Path

import pypff
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


class EmailProcessor:
    """
    EmailProcessor is a utility class for extracting, processing, and exporting
    email messages from PST files.

    This class provides methods to:
    - Load emails from a PST file using the pypff library.
    - Recursively extract messages from folders and subfolders within the PST file.
    - Format email delivery or creation times for use in filenames.
    - Save extracted emails as .eml files with sanitized filenames.
    - Export emails as PDF files using the reportlab library, including headers
      and message body.

    Attributes:
        file_path (str): Path to the PST file to be processed.
        emails (list): List of extracted email data dictionaries.

    Typical usage:
        processor = EmailProcessor("/path/to/file.pst")
        emails = processor.load_emails()
        processor.save_as_eml(email, "/output/dir")
        processor.save_as_pdf(email, "/output/dir")
    """

    def __init__(self, file_path):
        """
        Initialize the EmailProcessor with the path to a PST/OST file.

        Args:
            file_path (str): The file path to the PST/OST file to be processed.
        """
        self.file_path = file_path
        self.emails = []

    def extract_messages(self, folder, folder_path=""):
        """
        Recursively extract messages from the given folder and its subfolders.

        Args:
            folder: The pypff folder object to extract messages from.
            folder_path (str, optional): The current folder path for maintaining
                hierarchy. Defaults to "".
        """
        for i in range(folder.number_of_sub_messages):
            message = folder.get_sub_message(i)
            email_data = {"message": message, "folder_path": folder_path}
            self.emails.append(email_data)
        for j in range(folder.number_of_sub_folders):
            sub_folder = folder.get_sub_folder(j)
            sub_folder_name = sub_folder.name or "Unnamed Folder"
            self.extract_messages(
                sub_folder, os.path.join(folder_path, sub_folder_name)
            )

    def load_emails(self):
        """
        Load emails from the PST file and return them as a list of dictionaries.

        Returns:
            list: A list of dictionaries containing email data with 'message'
                and 'folder_path' keys.
        """
        pst_file = pypff.file()
        pst_file.open(self.file_path)

        root_folder = pst_file.get_root_folder()
        self.extract_messages(root_folder)
        pst_file.close()

        return self.emails

    def process_emails(self, output_dir, output_format="eml", verbose=False, dry_run=False):
        """
        Process emails with the given arguments and extract them to the specified formats.

        Args:
            output_dir (str): Output directory for extracted emails
            output_format (str): Output format ('eml', 'pdf', or 'both')
            verbose (bool): Enable verbose output
            dry_run (bool): Preview mode without saving files

        Returns:
            bool: True if processing was successful, False if errors occurred.

        This method performs the core email extraction workflow:
        - Validates the input PST/OST file
        - Creates the output directory structure
        - Loads and processes all emails from the PST file
        - Saves emails in the requested format(s)
        - Extracts and saves attachments
        - Provides progress feedback and error handling
        """
        from pst_processor import PSTProcessor
        from file_saver import FileSaver
        
        # Validate input file using PSTProcessor
        pst_processor = PSTProcessor()
        if not pst_processor.validate_pst_file(self.file_path):
            print(f"❌ Error: Input file '{self.file_path}' does not exist or is not a valid PST/OST file.")
            return False

        # Create output directory
        output_path = Path(output_dir)
        if not dry_run:
            output_path.mkdir(parents=True, exist_ok=True)

        print(f"\n🔍 Processing PST/OST file: {self.file_path}")
        print(f"📂 Output directory: {output_dir}")
        print(f"📄 Output format: {output_format}")
        print("-" * 50)

        try:
            # Load emails
            emails = self.load_emails()

            print(f"✅ Found {len(emails)} emails to process")

            if dry_run:
                print("\n📋 DRY RUN - Files that would be created:")
                for i, email_data in enumerate(emails[:10]):  # Show first 10 as preview
                    email = email_data["message"]
                    folder_path = email_data["folder_path"]

                    # Format delivery time for filename
                    date_prefix = self._format_delivery_time(email)
                    subject = email.subject or "no_subject"
                    clean_subject = subject[:50].strip()

                    if output_format in ["eml", "both"]:
                        print(f"  📧 {folder_path}/{date_prefix} - {clean_subject}.eml")
                    if output_format in ["pdf", "both"]:
                        print(f"  📄 {folder_path}/{date_prefix} - {clean_subject}.pdf")

                if len(emails) > 10:
                    print(f"  ... and {len(emails) - 10} more emails")
                return True

            # Initialize file saver
            file_saver = FileSaver(output_dir)

            # Process each email
            processed = 0
            for email_data in emails:
                email = email_data["message"]
                folder_path = email_data["folder_path"]

                try:
                    # Create folder structure using FileSaver
                    full_folder_path = file_saver._create_full_path(folder_path)

                    # Save email in requested format(s)
                    if output_format in ["eml", "both"]:
                        self.save_as_eml(email, full_folder_path)
                        if verbose:
                            print(
                                f"✅ Saved EML: {folder_path}/{email.subject or 'No Subject'}"
                            )

                    if output_format in ["pdf", "both"]:
                        self.save_as_pdf(email, full_folder_path)
                        if verbose:
                            print(
                                f"✅ Saved PDF: {folder_path}/{email.subject or 'No Subject'}"
                            )

                    # Save attachments if any
                    try:
                        num_attachments = email.number_of_attachments
                        if num_attachments and num_attachments > 0:
                            attachments_folder = os.path.join(full_folder_path, "attachments")
                            os.makedirs(attachments_folder, exist_ok=True)

                            for i in range(num_attachments):
                                try:
                                    attachment = email.get_attachment(i)
                                    file_saver.save_attachment(attachment, attachments_folder)
                                    if verbose:
                                        print(f"📎 Saved attachment: {attachment.name}")
                                except Exception as e:
                                    print(f"⚠️  Error saving attachment {i}: {e}")
                    except Exception:
                        # Silently assume no attachments if access fails
                        pass

                    processed += 1

                    # Progress indicator
                    if not verbose and processed % 50 == 0:
                        print(f"📧 Processed {processed}/{len(emails)} emails...")

                except Exception as e:
                    print(
                        f"❌ Error processing email '{email.subject or 'No Subject'}': {e}"
                    )
                    continue

            print(f"\n🎉 Successfully processed {processed}/{len(emails)} emails!")
            print(f"📂 Files saved to: {output_dir}")
            return True

        except Exception as e:
            print(f"❌ Error: {e}")
            if verbose:
                import traceback
                traceback.print_exc()
            return False

    def save_as_eml(self, email, output_path):
        """
        Save the email as an .eml file with sanitized filename including date prefix.

        Args:
            email: The pypff email object to save.
            output_path (str): The directory path where the .eml file will be saved.
        """
        # Format delivery time for filename
        date_prefix = self._format_delivery_time(email)
        subject = email.subject or "no_subject"
        # Remove invalid characters instead of replacing with dashes
        clean_subject = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", subject)
        file_name = f"{date_prefix} - {clean_subject[:50].strip()}.eml"
        full_path = os.path.join(output_path, file_name)

        # Extract recipient information from transport headers
        recipient_list = []
        transport_headers = getattr(email, "transport_headers", None)
        if transport_headers:
            try:
                msg = py_email.message_from_string(transport_headers, policy=policy.default)
                to_recipients = msg.get_all("To", [])
                cc_recipients = msg.get_all("Cc", [])
                bcc_recipients = msg.get_all("Bcc", [])
                all_recipients = to_recipients + cc_recipients + bcc_recipients
                recipient_list = [str(recipient) for recipient in all_recipients if recipient]
            except Exception:
                pass
        
        # Fallback to display_to if available
        if not recipient_list and hasattr(email, "display_to") and email.display_to:
            recipient_list = [email.display_to]
        
        # Final fallback
        if not recipient_list:
            recipient_list = ["Unknown Recipient"]

        # Get email body and decode if it's in bytes
        body = email.plain_text_body or email.html_body or ""
        if isinstance(body, bytes):
            # Try to detect encoding using multiple approaches
            detected_body = self._decode_email_body(body)
            body = detected_body

        with open(full_path, "w", encoding="utf-8") as eml_file:
            eml_file.write(f"Subject: {email.subject}\n")
            eml_file.write(f"From: {email.sender_name}\n")
            eml_file.write(f"To: {', '.join(recipient_list)}\n")
            eml_file.write("\n")
            eml_file.write(body)

    def save_as_pdf(self, email, output_path):
        """
        Save the email as a .pdf file using reportlab with headers and body content.

        Args:
            email: The pypff email object to save.
            output_path (str): The directory path where the .pdf file will be saved.
        """
        # Format delivery time for filename
        date_prefix = self._format_delivery_time(email)
        subject = email.subject or "no_subject"
        # Remove invalid characters instead of replacing with dashes
        clean_subject = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", subject)
        file_name = f"{date_prefix} - {clean_subject[:50].strip()}.pdf"
        full_path = os.path.join(output_path, file_name)

        # Use the email's plain text body if available, otherwise fall back to
        # HTML or a default message
        content = (
            email.plain_text_body
            or email.html_body
            or "No content available for this email."
        )

        # Decode the body if it is in bytes
        if isinstance(content, bytes):
            content = self._decode_email_body(content)

        # Create the PDF
        c = canvas.Canvas(full_path, pagesize=letter)
        c.setFont("Helvetica", 12)

        # Write email headers
        c.drawString(50, 750, f"Subject: {email.subject or 'No Subject'}")
        c.drawString(
            50,
            730,
            f"From: {email.sender_name or 'Unknown Sender'} <{email.sender_email_address or 'Unknown Email'}>",
        )
        
        # Extract recipient information from transport headers
        to_field = "Unknown Recipient"
        transport_headers = getattr(email, "transport_headers", None)
        if transport_headers:
            try:
                msg = py_email.message_from_string(transport_headers, policy=policy.default)
                to_recipients = msg.get_all("To", [])
                if to_recipients:
                    to_field = ", ".join(str(recipient) for recipient in to_recipients)
            except Exception:
                pass
        
        # Fallback to display_to if available
        if to_field == "Unknown Recipient" and hasattr(email, "display_to") and email.display_to:
            to_field = email.display_to
            
        c.drawString(50, 710, f"To: {to_field}")

        # Write email body
        y_position = 690
        line_height = 14
        for line in content.splitlines():
            if (
                y_position < 50
            ):  # Start a new page if the content exceeds the current page
                c.showPage()
                c.setFont("Helvetica", 12)
                y_position = 750
            c.drawString(50, y_position, line)
            y_position -= line_height

        # Save the PDF
        c.save()

    def _format_delivery_time(self, email):
        """
        Format the delivery time for use in filenames as [YYYY-MM-DD].

        Falls back to creation_time or current date if delivery_time is not
        available.

        Args:
            email: The email object containing delivery_time and creation_time
                attributes.

        Returns:
            str: Formatted date string in the format [YYYY-MM-DD].
        """
        try:
            # Try to get delivery_time first
            delivery_time = getattr(email, "delivery_time", None)
            if delivery_time:
                return f"[{delivery_time.strftime('%Y-%m-%d')}]"

            # Fall back to creation_time
            creation_time = getattr(email, "creation_time", None)
            if creation_time:
                return f"[{creation_time.strftime('%Y-%m-%d')}]"

            # Last resort: use current date
            return f"[{datetime.now().strftime('%Y-%m-%d')}]"
        except (AttributeError, TypeError):
            # If there's any error with date formatting, use current date
            return f"[{datetime.now().strftime('%Y-%m-%d')}]"

    def _decode_email_body(self, body_bytes):
        """
        Decode email body bytes with smart encoding detection.

        This method tries multiple approaches to detect and decode the correct
        encoding for email body content, handling common encoding issues.

        Args:
            body_bytes (bytes): The raw email body content in bytes.

        Returns:
            str: The decoded email body as a string.
        """
        if not isinstance(body_bytes, bytes):
            return str(body_bytes)

        # List of encodings to try, in order of preference
        # Based on common email encodings and Portuguese/international content
        encodings_to_try = [
            'utf-8',           # Most common modern encoding
            'iso-8859-1',      # Latin-1, very common for European languages
            'windows-1252',    # Windows Latin-1, common in Windows emails
            'cp1252',          # Alternative name for windows-1252
            'iso-8859-15',     # Latin-9, includes Euro symbol
            'utf-16',          # Unicode with BOM
            'utf-16le',        # Little-endian UTF-16
            'utf-16be',        # Big-endian UTF-16
            'ascii',           # Basic ASCII as last resort
        ]

        # First, try to detect if it's a double-encoded UTF-8 (common issue)
        # This happens when UTF-8 encoded text is incorrectly decoded as Latin-1
        # and then re-encoded as UTF-8
        try:
            # Try decoding as Latin-1 first, then re-encode and decode as UTF-8
            temp_decoded = body_bytes.decode('latin-1')
            if any(char in temp_decoded for char in ['Ã§', 'Ã£', 'Ã¡', 'Ã©', 'Ã­', 'Ã³', 'Ãº']):
                # Likely double-encoded UTF-8, try to fix it
                try:
                    fixed_bytes = temp_decoded.encode('latin-1')
                    return fixed_bytes.decode('utf-8')
                except (UnicodeDecodeError, UnicodeEncodeError):
                    pass
        except UnicodeDecodeError:
            pass

        # Try each encoding in order
        for encoding in encodings_to_try:
            try:
                decoded = body_bytes.decode(encoding)
                
                # Validate the result - check for common problematic patterns
                if encoding == 'utf-8' and any(char in decoded for char in ['Ã§', 'Ã£', 'Ã¡']):
                    # This might be wrongly decoded UTF-8, skip to next encoding
                    continue
                
                # If we get here, the decoding worked
                return decoded
                
            except (UnicodeDecodeError, UnicodeError):
                continue

        # If all encodings fail, use UTF-8 with error replacement as last resort
        try:
            return body_bytes.decode('utf-8', errors='replace')
        except Exception:
            # Absolute last resort
            return f"Error: Could not decode email body (length: {len(body_bytes)} bytes)"
