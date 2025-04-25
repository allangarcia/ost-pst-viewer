# main.py

import os
import logging
from email_processor import EmailProcessor
from file_saver import FileSaver
from folder_structure import FolderStructure

# Configure logging
logging.basicConfig(
    filename='error.log',  # Log file name
    level=logging.ERROR,   # Log level
    format='%(asctime)s - %(levelname)s - %(message)s'  # Log format
)

def main():
    print("Welcome to the OST/PST Viewer")
    
    # Initialize components
    base_directory = input("Enter the base directory to save files: ")
    file_saver = FileSaver(base_directory=base_directory)
    folder_structure = FolderStructure(base_path=base_directory)

    # Load emails from a specified OST/PST file
    ost_pst_file = input("Enter the path to the OST/PST file: ")
    email_processor = EmailProcessor(file_path=ost_pst_file)
    emails = email_processor.load_emails()

    # Ask for the save format once
    save_format = input("Save all emails as (eml/pdf): ").strip().lower()
    if save_format not in ['eml', 'pdf']:
        print("Invalid format. Exiting.")
        return

    # Process and save emails
    for email_data in emails:
        email = email_data["message"]
        folder_path = folder_structure.create_folder(email_data["folder_path"])
        
        # Save email in the chosen format
        if save_format == 'eml':
            file_saver.save_email(email, folder_path, format='eml')
        elif save_format == 'pdf':
            file_saver.save_email(email, folder_path, format='pdf')

        # Save attachments
        try:
            # Attempt to retrieve the number of attachments
            try:
                num_attachments = email.number_of_attachments
            except OSError as e:
                logging.error(f"Error retrieving number of attachments for email '{email.subject}': {e}")
                num_attachments = 0  # Assume no attachments if retrieval fails

            if num_attachments > 0:  # Check if there are attachments
                for i in range(num_attachments):
                    try:
                        # Attempt to retrieve the attachment
                        attachment = email.get_attachment(i)
                        if attachment is None:
                            logging.error(f"Attachment {i + 1} for email '{email.subject}' is missing or invalid. Skipping.")
                            continue

                        # Save the attachment
                        file_saver.save_attachment(attachment, folder_path)
                    except OSError as e:
                        logging.error(f"Error saving attachment {i + 1} for email '{email.subject}': {e}")
        except Exception as e:
            logging.error(f"Unexpected error processing attachments for email '{email.subject}': {e}")

    print("Email processing complete.")

if __name__ == "__main__":
    main()