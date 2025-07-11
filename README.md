# PST/OST Email Exporter

This project is a PST/OST email exporter that allows users to extract Outlook emails from PST/OST files, save them in .eml or .pdf formats, and save attachments in their original formats while maintaining the folder structure. The application does not require the original Outlook password to function.

## Features

- 🔍 Load emails from PST/OST files without requiring passwords
- 📧 Save emails as .eml or .pdf files with date prefixes
- 📎 Save attachments in their original format
- 📁 Maintain the original folder structure of emails
- 🖥️ Interactive mode for easy file selection and configuration
- ⚡ Command-line interface for batch processing
- 🔧 Smart filename sanitization and duplicate handling
- 📅 Automatic date extraction for chronological organization

## Project Structure

```
pst-exporter/
├── pst-exporter.py             # Main script (interactive + CLI modes)
├── pst-exporter.bat            # Windows wrapper for easy execution
├── src/                        # Core application code
│   ├── email_processor.py      # Handles email extraction and processing
│   └── file_saver.py           # Responsible for saving emails and attachments
├── tests/                      # Unit tests
│   ├── test_email_processor.py # Unit tests for EmailProcessor
│   └── test_file_saver.py      # Unit tests for FileSaver
├── pst_files/                  # Place your PST/OST files here
├── output/                     # Generated output files
├── requirements.txt            # Project dependencies
├── .gitignore                  # Git ignore rules
└── README.md                   # Project documentation
```

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd pst-exporter/
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### 🎯 Quick Start (Interactive Mode)

Simply run the script without any arguments to enter interactive mode:

```bash
# Linux/Mac/Codespaces
./pst-exporter.py

# Windows
pst-exporter.bat
```

The interactive mode will:

- 📁 Auto-discover PST/OST files in the `pst_files/` directory
- 🎮 Guide you through file selection and format options
- 🎯 Suggest smart output directory names
- ✨ Provide a user-friendly experience

### ⚡ Command Line Mode

For batch processing or automation:

```bash
# Basic usage
./pst-exporter.py -i pst_files/archive.pst

# Specify output format and directory
./pst-exporter.py -i pst_files/archive.pst -f pdf -o extracted_emails

# Both EML and PDF with verbose output
./pst-exporter.py -i pst_files/archive.pst -f both -v

# Dry run to preview what would be processed
./pst-exporter.py -i pst_files/archive.pst --dry-run
```

### 📋 Command Line Options

| Option | Description | Example |
|--------|-------------|---------|
| `-i`, `--input` | PST/OST file path | `-i archive.pst` |
| `-o`, `--output` | Output directory | `-o extracted_emails` |
| `-f`, `--format` | Output format (eml/pdf/both) | `-f both` |
| `-v`, `--verbose` | Verbose output | `-v` |
| `--dry-run` | Preview mode | `--dry-run` |
| `--interactive` | Force interactive mode | `--interactive` |

### 📂 File Organization

Output files are organized with date prefixes for easy chronological sorting:

```
output/
├── [2024-03-15] - Meeting Notes.eml
├── [2024-03-15] - Project Update.pdf
├── Inbox/
│   ├── [2024-03-16] - Important Email.eml
│   └── attachments/
│       └── document.pdf
└── Sent Items/
    └── [2024-03-17] - Reply.eml
```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## Desirable changes

- Delegate more methods from the main pst-exporter.py script to other files
- Separate some methods by subject to be in their own files
- Separete the command-line logic from the main pst-exporter.py

## License

This project is licensed under the MIT License. See the LICENSE file for details.