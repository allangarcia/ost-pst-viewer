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
- 🌐 Smart encoding detection for international characters
- 📬 Improved recipient handling from email headers
- ✅ Comprehensive error handling and validation

## Project Structure

```text
pst-exporter/
├── pst-exporter.py             # Main script (interactive + CLI modes)
├── pst-exporter.bat            # Windows convenience wrapper
├── src/                        # Core application code
│   ├── email_processor.py      # Handles email extraction and processing
│   ├── file_saver.py           # Responsible for saving emails and attachments
│   └── pst_processor.py        # Handles PST file discovery and validation
├── tests/                      # Unit tests
│   ├── test_email_processor.py # Unit tests for EmailProcessor
│   ├── test_file_saver.py      # Unit tests for FileSaver
│   └── test_pst_processor.py   # Unit tests for PSTProcessor
├── pst_files/                  # Place your PST/OST files here
├── output/                     # Generated output files
├── .github/                    # GitHub workflows and templates
├── .vscode/                    # VS Code workspace settings
├── requirements.txt            # Project dependencies
├── pyproject.toml              # Python project configuration
├── setup.cfg                   # Additional project setup
├── .gitignore                  # Git ignore rules
├── SECURITY.md                 # Security policy
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

3. **For Windows users**: The included `pst-exporter.bat` file provides a convenient way to run the application without needing to type Python commands.

## Usage

### 🎯 Quick Start (Interactive Mode)

Simply run the script without any arguments to enter interactive mode:

```bash
# Linux/Mac/Codespaces
./pst-exporter.py

# Windows (using batch file - recommended)
pst-exporter.bat

# Windows (using Python directly)
python pst-exporter.py
```

The interactive mode will:

- 📁 Auto-discover PST/OST files in the `pst_files/` directory
- 🎮 Guide you through file selection and format options
- 🎯 Suggest smart output directory names
- ✨ Provide a user-friendly experience

### ⚡ Command Line Mode

For batch processing or automation:

```bash
# Basic usage (Linux/Mac)
./pst-exporter.py -i pst_files/archive.pst

# Basic usage (Windows - using batch file)
pst-exporter.bat -i pst_files/archive.pst

# Specify output format and directory
./pst-exporter.py -i pst_files/archive.pst -f pdf -o extracted_emails

# Both EML and PDF with verbose output
./pst-exporter.py -i pst_files/archive.pst -f both -v

# Dry run to preview what would be processed
./pst-exporter.py -i pst_files/archive.pst --dry-run
```

**Note for Windows users**: Replace `./pst-exporter.py` with `pst-exporter.bat` in any of the above commands for easier execution.

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

```text
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

## Recent Improvements

Recent updates to the project include:

- ✅ **Refactored architecture**: Email processing logic moved to EmailProcessor class for better organization
- ✅ **PST file operations**: Extracted into dedicated PSTProcessor class for improved modularity
- ✅ **Enhanced encoding detection**: Smart handling of international characters and email encodings
- ✅ **Better recipient handling**: Improved extraction of recipient information from email headers
- ✅ **Comprehensive testing**: Unit tests for all major components

## Future Enhancements

Potential areas for improvement:

- Separate command-line argument parsing into its own module
- Add configuration file support for default settings and export formatting
- Implement email filtering and search capabilities
- Add support for additional output formats (HTML, JSON)
- Make a graphical interface using Electron or similar
- Improve the test files for better understanding and test automation on PR

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

This is a fork. So you can submit your ideas to the original author as well.

## License

This project is licensed under the MIT License.
