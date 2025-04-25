import pypff
import os
import pdfkit
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

class EmailProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.emails = []

    def extract_messages(self, folder, folder_path=""):
        """
        Recursively extracts messages from the given folder and its subfolders.
        """
        for i in range(folder.number_of_sub_messages):
            message = folder.get_sub_message(i)
            email_data = {
                "message": message,
                "folder_path": folder_path
            }
            self.emails.append(email_data)
        for j in range(folder.number_of_sub_folders):
            sub_folder = folder.get_sub_folder(j)
            sub_folder_name = sub_folder.name or "Unnamed Folder"
            self.extract_messages(sub_folder, os.path.join(folder_path, sub_folder_name))

    def load_emails(self):
        """
        Loads emails from the PST file and returns them as a list of dictionaries.
        """
        pst_file = pypff.file()
        pst_file.open(self.file_path)

        root_folder = pst_file.get_root_folder()
        self.extract_messages(root_folder)
        pst_file.close()

        return self.emails

    def save_as_eml(self, email, output_path):
        subject = email.subject or "no_subject"
        file_name = f"{subject[:50].strip().replace('/', '-')}.eml"
        full_path = os.path.join(output_path, file_name)

        with open(full_path, "w", encoding="utf-8") as eml_file:
            eml_file.write(f"Subject: {email.subject}\n")
            eml_file.write(f"From: {email.sender_name}\n")
            eml_file.write(f"To: {email.display_to}\n")
            eml_file.write("\n")
            eml_file.write(email.plain_text_body or email.html_body or "")

    def save_as_pdf(self, email, output_path):
        """
        Saves the email as a .pdf file using reportlab.
        """
        subject = email.subject or "no_subject"
        file_name = f"{subject[:50].strip().replace('/', '-')}.pdf"
        full_path = os.path.join(output_path, file_name)

        # Use the email's plain text body if available, otherwise fall back to HTML or a default message
        content = email.plain_text_body or email.html_body or "No content available for this email."

        # Decode the body if it is in bytes
        if isinstance(content, bytes):
            content = content.decode('utf-8', errors='replace')

        # Create the PDF
        c = canvas.Canvas(full_path, pagesize=letter)
        c.setFont("Helvetica", 12)

        # Write email headers
        c.drawString(50, 750, f"Subject: {email.subject or 'No Subject'}")
        c.drawString(50, 730, f"From: {email.sender_name or 'Unknown Sender'} <{email.sender_email_address or 'Unknown Email'}>")
        c.drawString(50, 710, f"To: {email.display_to or 'Unknown Recipient'}")

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