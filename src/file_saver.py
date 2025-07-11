import email as py_email
import os
import re
from email import policy

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


class FileSaver:
    """
    FileSaver handles saving emails and attachments to disk with smart filename
    handling.

    This class provides methods to:
    - Save emails in EML or PDF format with sanitized filenames
    - Save attachments in their original format
    - Create unique filenames to avoid duplicates
    - Manage folder structure creation

    Attributes:
        base_directory (str): Base directory where all files will be saved.
    """

    def __init__(self, base_directory):
        """
        Initialize the FileSaver with a base directory.

        Args:
            base_directory (str): The base directory where files will be saved.
        """
        self.base_directory = base_directory

    def save_email(self, email, folder_path, output_format="eml"):
        """
        Save the email in the specified format while maintaining folder structure.

        Args:
            email: The pypff email object to save.
            folder_path (str): The relative folder path where the email should be
                saved.
            output_format (str, optional): The output format ('eml' or 'pdf').
                Defaults to 'eml'.
        """
        folder_path = self._create_full_path(folder_path)
        if output_format == "eml":
            self._save_as_eml(email, folder_path)
        elif output_format == "pdf":
            self._save_as_pdf(email, folder_path)

    def save_attachment(self, attachment, folder_path):
        """
        Save an attachment to the specified folder with unique filename handling.

        Args:
            attachment: The pypff attachment object to save.
            folder_path (str): The directory path where the attachment will be
                saved.

        Note:
            If the attachment has no name, a default name will be generated.
            Duplicate filenames will be handled by appending a counter
            (e.g., file_1.ext).
        """
        try:
            # Get the attachment name and sanitize it
            raw_name = attachment.name or f"attachment_{id(attachment)}"
            sanitized_name = self._sanitize_filename(raw_name)

            # Split name and extension
            if "." in sanitized_name:
                base_name, extension = sanitized_name.rsplit(".", 1)
            else:
                base_name = sanitized_name
                extension = "bin"  # Default extension for files without one

            # Get unique filename with counter if needed
            attachment_path = self._get_unique_filename(
                folder_path, base_name, extension
            )

            # Check if attachment data is available
            attachment_data = attachment.read_buffer()
            if not attachment_data:
                print(f"Warning: No data available for attachment '{attachment.name}'")
                return

            # Write the attachment content to a file
            with open(attachment_path, "wb") as attachment_file:
                attachment_file.write(attachment_data)
        except PermissionError as e:
            print(f"Permission denied saving attachment '{attachment.name}': {e}")
        except OSError as e:
            print(f"OS error saving attachment '{attachment.name}': {e}")
        except Exception as e:
            print(f"Error saving attachment '{attachment.name}': {e}")

    def _sanitize_filename(self, filename, max_length=200):
        """
        Sanitize a filename by removing invalid characters and limiting length.

        Args:
            filename (str): The original filename to sanitize.
            max_length (int, optional): Maximum length for the filename.
                Defaults to 200.

        Returns:
            str: A sanitized filename safe for filesystem use.
        """
        if not filename:
            return "unnamed_file"

        # Remove invalid characters for Windows and Unix
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        sanitized = re.sub(invalid_chars, "", filename)

        # Remove leading/trailing spaces and dots
        sanitized = sanitized.strip(" .")

        # Limit length
        if len(sanitized) > max_length:
            sanitized = sanitized[:max_length]

        # Ensure it's not empty after sanitization
        if not sanitized:
            sanitized = "unnamed_file"

        return sanitized

    def _get_unique_filename(self, folder_path, base_name, extension):
        """
        Generate a unique filename by appending a counter if the file already exists.

        Returns the full file path with a unique name.

        Args:
            folder_path (str): The directory where the file will be saved.
            base_name (str): The base filename without extension.
            extension (str): The file extension without the dot.

        Returns:
            str: Full file path with a unique filename (e.g., file_1.ext,
                file_2.ext).
        """
        file_path = os.path.join(folder_path, f"{base_name}.{extension}")

        if not os.path.exists(file_path):
            return file_path

        counter = 1
        while True:
            new_filename = f"{base_name}_{counter}.{extension}"
            file_path = os.path.join(folder_path, new_filename)
            if not os.path.exists(file_path):
                return file_path
            counter += 1

    def _create_full_path(self, folder_path):
        """
        Create the full path for the folder where the email or attachment will be saved.

        Automatically creates any missing directories in the path.

        Args:
            folder_path (str): The relative folder path to create within the
                base directory.

        Returns:
            str: The full absolute path to the created directory.

        Raises:
            PermissionError: If there are insufficient permissions to create
                the directory.
            OSError: If there's a filesystem error preventing directory
                creation.
        """
        try:
            full_path = os.path.join(self.base_directory, folder_path)
            os.makedirs(full_path, exist_ok=True)
            return full_path
        except PermissionError as e:
            print(f"Permission denied creating directory '{full_path}': {e}")
            raise
        except OSError as e:
            print(f"OS error creating directory '{full_path}': {e}")
            raise

    def _save_as_eml(self, email, folder_path):
        """
        Saves the email as an .eml file with headers and body content.

        Args:
            email: The pypff email object to save.
            folder_path (str): The directory path where the .eml file will be saved.

        Note:
            - Attempts to extract proper email headers from transport_headers
            - Handles various text encodings (UTF-8, Latin-1, CP1252, ASCII)
            - Uses unique filename generation to avoid conflicts
            - Falls back gracefully when headers cannot be parsed
        """
        try:
            # Sanitize the filename and get unique path
            safe_subject = self._sanitize_filename(email.subject or "No Subject")
            eml_file_path = self._get_unique_filename(folder_path, safe_subject, "eml")

            # Extract email components
            headers = getattr(email, "transport_headers", None)
            sender_email = "Unknown Email"
            sender = "Unknown Sender"
            recipient_list = []

            if headers:
                try:
                    msg = py_email.message_from_string(headers, policy=policy.default)
                    sender = msg.get("From", "Unknown Sender")
                    sender_email_match = re.search(r"<(.+?)>", sender)
                    if sender_email_match:
                        sender_email = sender_email_match.group(1)

                    to_recipients = msg.get_all("To", [])
                    cc_recipients = msg.get_all("Cc", [])
                    bcc_recipients = msg.get_all("Bcc", [])
                    all_recipients = to_recipients + cc_recipients + bcc_recipients
                    for recipient in all_recipients:
                        recipient_list.append(recipient)
                except Exception as e:
                    print(f"Warning: Error parsing email headers: {e}")

            body = email.plain_text_body or email.html_body or b"No content available."

            # Decode the body if it is in bytes
            if isinstance(body, bytes):
                try:
                    body = body.decode("utf-8", errors="replace")
                except UnicodeDecodeError:
                    # Try other common encodings
                    for encoding in ["latin-1", "cp1252", "ascii"]:
                        try:
                            body = body.decode(encoding, errors="replace")
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        body = "Error: Could not decode email body"

            # Write the email to an .eml file
            with open(eml_file_path, "w", encoding="utf-8") as eml_file:
                # Write headers
                eml_file.write(f"Subject: {email.subject or 'No Subject'}\n")
                eml_file.write(f"From: {sender} <{sender_email}>\n")
                eml_file.write("To: ")
                eml_file.write(", ".join(recipient_list))
                eml_file.write("\n\n")  # Separate headers from the body

                # Write body
                eml_file.write(body)

        except PermissionError as e:
            print(f"Permission denied saving email '{email.subject}': {e}")
        except OSError as e:
            print(f"OS error saving email '{email.subject}': {e}")
        except Exception as e:
            print(f"Error saving email '{email.subject}': {e}")

    def _save_as_pdf(self, email, folder_path):
        """
        Saves the email as a .pdf file using reportlab with headers and body content.

        Args:
            email: The pypff email object to save.
            folder_path (str): The directory path where the .pdf file will be saved.

        Note:
            - Creates a PDF with properly formatted email headers (Subject, From, To)
            - Handles multiple pages when content exceeds page boundaries
            - Attempts to extract recipient information from transport headers
            - Supports various text encodings for email body content
            - Uses unique filename generation to avoid conflicts
        """
        try:
            # Sanitize the filename and get unique path
            safe_subject = self._sanitize_filename(email.subject or "No Subject")
            pdf_file_path = self._get_unique_filename(folder_path, safe_subject, "pdf")

            # Use the email's plain text body if available, otherwise fall back
            # to HTML or a default message
            content = (
                email.plain_text_body
                or email.html_body
                or "No content available for this email."
            )

            # Decode the body if it is in bytes
            if isinstance(content, bytes):
                try:
                    content = content.decode("utf-8", errors="replace")
                except UnicodeDecodeError:
                    # Try other common encodings
                    for encoding in ["latin-1", "cp1252", "ascii"]:
                        try:
                            content = content.decode(encoding, errors="replace")
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        content = "Error: Could not decode email content"

            # Create the PDF
            c = canvas.Canvas(pdf_file_path, pagesize=letter)
            c.setFont("Helvetica", 12)

            # Write email headers
            c.drawString(50, 750, f"Subject: {email.subject or 'No Subject'}")
            sender = getattr(email, "sender_name", "Unknown Sender")
            sender_email = "Unknown Email"
            transport_headers = getattr(email, "transport_headers", None)
            if transport_headers:
                try:
                    msg = py_email.message_from_string(
                        transport_headers, policy=policy.default
                    )
                    sender_field = msg.get("From", "Unknown Sender")
                    sender_email_match = re.search(r"<(.+?)>", sender_field)
                    if sender_email_match:
                        sender_email = sender_email_match.group(1)
                        sender = sender_field
                except Exception:
                    pass
            c.drawString(50, 730, f"From: {sender} <{sender_email}>")
            to_field = "Unknown Recipient"
            if transport_headers:
                try:
                    msg = py_email.message_from_string(
                        transport_headers, policy=policy.default
                    )
                    to_recipients = msg.get_all("To", [])
                    if to_recipients:
                        to_field = ", ".join(to_recipients)
                except Exception:
                    pass
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

        except PermissionError as e:
            print(f"Permission denied saving PDF '{email.subject}': {e}")
        except OSError as e:
            print(f"OS error saving PDF '{email.subject}': {e}")
        except Exception as e:
            print(f"Error saving PDF '{email.subject}': {e}")

    def _save_original_attachment(self, attachment, folder_path):
        """
        Saves the attachment in its original format with unique filename handling.

        Args:
            attachment: The email attachment object to save.
            folder_path (str): The directory path where the attachment will be saved.

        Note:
            - Preserves the original file format and extension
            - Uses unique filename generation to avoid conflicts
            - Falls back to .bin extension for files without extensions
            - Handles various error conditions gracefully

        This method is an alternative to save_attachment() for cases where
        the attachment object has a different interface.
        """
        try:
            # Sanitize filename and get unique path
            raw_name = attachment.filename or f"attachment_{id(attachment)}"
            sanitized_name = self._sanitize_filename(raw_name)

            # Split name and extension
            if "." in sanitized_name:
                base_name, extension = sanitized_name.rsplit(".", 1)
            else:
                base_name = sanitized_name
                extension = "bin"

            # Get unique filename with counter if needed
            attachment_file_path = self._get_unique_filename(
                folder_path, base_name, extension
            )

            with open(attachment_file_path, "wb") as attachment_file:
                attachment_file.write(attachment.content)

        except PermissionError as e:
            filename = attachment.filename
            print(f"Permission denied saving original attachment '{filename}': {e}")
        except OSError as e:
            print(f"OS error saving original attachment '{attachment.filename}': {e}")
        except Exception as e:
            print(f"Error saving original attachment '{attachment.filename}': {e}")
