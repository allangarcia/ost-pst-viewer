import email as py_email
import os
import re
from datetime import datetime
from email import policy

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

        with open(full_path, "w", encoding="utf-8") as eml_file:
            eml_file.write(f"Subject: {email.subject}\n")
            eml_file.write(f"From: {email.sender_name}\n")
            eml_file.write(f"To: {', '.join(recipient_list)}\n")
            eml_file.write("\n")
            eml_file.write(email.plain_text_body or email.html_body or "")

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
            content = content.decode("utf-8", errors="replace")

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
