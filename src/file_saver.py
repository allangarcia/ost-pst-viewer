import os 
import pdfkit
import email as py_email
from email import policy
import re
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

class FileSaver:
    def __init__(self, base_directory):
        self.base_directory = base_directory

    def save_email(self, email, folder_path, format='eml'):
        """
        Saves the email in the specified format (eml or pdf) while maintaining the folder structure.
        """
        folder_path = self._create_full_path(folder_path)
        if format == 'eml':
            self._save_as_eml(email, folder_path)
        elif format == 'pdf':
            self._save_as_pdf(email, folder_path)

    def save_attachment(self, attachment, folder_path):
        """
        Saves an attachment to the specified folder.
        """
        try:
            attachment_name = attachment.name or f"attachment_{id(attachment)}"
            attachment_path = os.path.join(folder_path, attachment_name)

            # Write the attachment content to a file
            with open(attachment_path, 'wb') as attachment_file:
                attachment_file.write(attachment.read_buffer())
        except Exception as e:
            print(f"Error saving attachment '{attachment.name}': {e}")

    def _create_full_path(self, folder_path):
        """
        Creates the full path for the folder where the email or attachment will be saved.
        """
        full_path = os.path.join(self.base_directory, folder_path)
        os.makedirs(full_path, exist_ok=True)
        return full_path

    def _save_as_eml(self, email, folder_path):
        """
        Saves the email as an .eml file.
        """
        eml_file_path = os.path.join(folder_path, f"{email.subject or 'No Subject'}.eml")

        # Extract email components
        headers = getattr(email, 'transport_headers', None)
        sender_email = "Unknown Email"
        sender = "Unknown Sender"
        recipient_list = []

        if headers:
            try:
                msg = py_email.message_from_string(headers, policy=policy.default)
                sender = msg.get('From', "Unknown Sender")
                sender_email_match = re.search(r'<(.+?)>', sender)
                if sender_email_match:
                    sender_email = sender_email_match.group(1)

                to_recipients = msg.get_all('To', [])
                cc_recipients = msg.get_all('Cc', [])
                bcc_recipients = msg.get_all('Bcc', [])
                all_recipients = to_recipients + cc_recipients + bcc_recipients
                for recipient in all_recipients:
                    recipient_list.append(recipient)
            except Exception:
                pass

        body = email.plain_text_body or email.html_body or b"No content available."

        # Decode the body if it is in bytes
        if isinstance(body, bytes):
            body = body.decode('utf-8', errors='replace')

        # Write the email to an .eml file
        with open(eml_file_path, 'w', encoding='utf-8') as eml_file:
            # Write headers
            eml_file.write(f"Subject: {email.subject or 'No Subject'}\n")
            eml_file.write(f"From: {sender} <{sender_email}>\n")
            eml_file.write("To: ")
            eml_file.write(", ".join(recipient_list))
            eml_file.write("\n\n")  # Separate headers from the body

            # Write body
            eml_file.write(body)

    def _save_as_pdf(self, email, folder_path):
        """
        Saves the email as a .pdf file using reportlab.
        """
        pdf_file_path = os.path.join(folder_path, f"{email.subject or 'No Subject'}.pdf")

        # Use the email's plain text body if available, otherwise fall back to HTML or a default message
        content = email.plain_text_body or email.html_body or "No content available for this email."

        # Decode the body if it is in bytes
        if isinstance(content, bytes):
            content = content.decode('utf-8', errors='replace')

        # Create the PDF
        c = canvas.Canvas(pdf_file_path, pagesize=letter)
        c.setFont("Helvetica", 12)

        # Write email headers
        c.drawString(50, 750, f"Subject: {email.subject or 'No Subject'}")
        sender = getattr(email, 'sender_name', 'Unknown Sender')
        sender_email = 'Unknown Email'
        transport_headers = getattr(email, 'transport_headers', None)
        if transport_headers:
            try:
                msg = py_email.message_from_string(transport_headers, policy=policy.default)
                sender_field = msg.get('From', 'Unknown Sender')
                sender_email_match = re.search(r'<(.+?)>', sender_field)
                if sender_email_match:
                    sender_email = sender_email_match.group(1)
                    sender = sender_field
            except Exception:
                pass
        c.drawString(50, 730, f"From: {sender} <{sender_email}>")
        to_field = 'Unknown Recipient'
        if transport_headers:
            try:
                msg = py_email.message_from_string(transport_headers, policy=policy.default)
                to_recipients = msg.get_all('To', [])
                if to_recipients:
                    to_field = ", ".join(to_recipients)
            except Exception:
                pass
        c.drawString(50, 710, f"To: {to_field}")

        # Write email body
        y_position = 690
        line_height = 14
        for line in content.splitlines():
            if y_position < 50:  # Start a new page if the content exceeds the current page
                c.showPage()
                c.setFont("Helvetica", 12)
                y_position = 750
            c.drawString(50, y_position, line)
            y_position -= line_height

        # Save the PDF
        c.save()

    def _save_original_attachment(self, attachment, folder_path):
        """
        Saves the attachment in its original format.
        """
        attachment_file_path = os.path.join(folder_path, attachment.filename)
        with open(attachment_file_path, 'wb') as attachment_file:
            attachment_file.write(attachment.content)