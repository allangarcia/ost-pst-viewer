# OST/PST Viewer

This project is an OST/PST viewer that allows users to open Outlook emails, save them in .eml or .pdf formats, and save attachments in their original formats while maintaining the folder structure. The application does not require the old Outlook password to function.

## Features

- Load emails from OST/PST files.
- Save emails as .eml or .pdf files.
- Save attachments in their original format.
- Maintain the original folder structure of emails.

## Project Structure

```
ost-pst-viewer
├── src
│   ├── main.py               # Entry point of the application
│   ├── email_processor.py     # Handles email extraction and saving
│   ├── file_saver.py          # Responsible for saving emails and attachments
│   ├── folder_structure.py      # Manages folder structure creation
│   └── utils
│       └── helpers.py         # Utility functions for the project
├── tests
│   ├── test_email_processor.py # Unit tests for EmailProcessor
│   ├── test_file_saver.py      # Unit tests for FileSaver
│   ├── test_folder_structure.py  # Unit tests for FolderStructure
│   └── test_helpers.py         # Unit tests for utility functions
├── requirements.txt            # Project dependencies
├── .gitignore                  # Files and directories to ignore in version control
└── README.md                   # Project documentation
```

## Installation

1. Clone the repository:
   ```
   git clone <repository-url>
   ```
2. Navigate to the project directory:
   ```
   cd ost-pst-viewer
   ```
3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

To run the application, execute the following command:
```
python src/main.py
```

Follow the on-screen instructions to load OST/PST files and save emails and attachments.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.