#!/usr/bin/env python3
"""
PST/OST Email Exporter - Main Script
This script processes PST/OST files and extracts emails to various formats.
"""

import argparse
import glob
import os
import sys
from pathlib import Path

# Add src directory to path to import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from email_processor import EmailProcessor  # noqa: E402
from file_saver import FileSaver  # noqa: E402


def find_pst_files():
    """
    Find all PST/OST files in the pst_files directory.

    Returns:
        list: A sorted list of file paths to PST/OST files found in the
            pst_files directory. Returns an empty list if no files are found
            or if the directory doesn't exist.

    Note:
        Creates the pst_files directory if it doesn't exist.
    """
    pst_dir = "pst_files"
    if not os.path.exists(pst_dir):
        os.makedirs(pst_dir, exist_ok=True)
        return []

    pst_files = glob.glob(os.path.join(pst_dir, "*.pst")) + glob.glob(
        os.path.join(pst_dir, "*.ost")
    )
    return sorted(pst_files)


def interactive_mode():
    """
    Run the application in interactive mode with user prompts.

    Returns:
        dict or None: A dictionary containing user-selected options with keys:
                     'input', 'output', 'format', 'verbose', 'dry_run'.
                     Returns None if the user cancels or if no valid files are found.

    This function provides a user-friendly interface that:
    - Automatically finds PST/OST files in the pst_files directory
    - Allows users to select files or specify custom paths
    - Prompts for output format preferences (EML, PDF, or both)
    - Allows customization of output directory
    - Offers verbose output option
    """
    print("🔍 PST/OST Email Exporter - Interactive Mode")
    print("=" * 50)

    # Find PST files
    pst_files = find_pst_files()

    if not pst_files:
        print("⚠️  No PST/OST files found in 'pst_files' directory")
        print(
            "Please place your PST/OST files in the 'pst_files' directory and run again."
        )

        # Ask if user wants to specify custom path
        custom = (
            input("\nWould you like to specify a custom path? (y/n): ").lower().strip()
        )
        if custom in ["y", "yes"]:
            pst_file = input("Enter the full path to your PST/OST file: ").strip()
            if not os.path.exists(pst_file):
                print(f"❌ File '{pst_file}' does not exist")
                return None
        else:
            return None
    else:
        print("📁 Found PST/OST files:")
        for i, file in enumerate(pst_files, 1):
            print(f"  [{i}] {file}")

        print("  [c] Custom path")

        while True:
            choice = input(
                f"\nSelect a file (1-{len(pst_files)}) or 'c' for custom path: "
            ).strip()

            if choice.lower() == "c":
                pst_file = input("Enter the full path to your PST/OST file: ").strip()
                if not os.path.exists(pst_file):
                    print(f"❌ File '{pst_file}' does not exist")
                    continue
                break

            try:
                file_num = int(choice)
                if 1 <= file_num <= len(pst_files):
                    pst_file = pst_files[file_num - 1]
                    break
                else:
                    print(
                        f"❌ Invalid choice. Please enter a number between 1 and {len(pst_files)}"
                    )
            except ValueError:
                print("❌ Invalid input. Please enter a number or 'c'")

    # Ask for output format
    print("\n📄 Select output format:")
    print("  [1] EML files only")
    print("  [2] PDF files only")
    print("  [3] Both EML and PDF")

    while True:
        format_choice = input("Enter your choice (1-3): ").strip()
        if format_choice == "1":
            output_format = "eml"
            break
        elif format_choice == "2":
            output_format = "pdf"
            break
        elif format_choice == "3":
            output_format = "both"
            break
        else:
            print("❌ Invalid choice. Please enter 1, 2, or 3")

    # Ask for output directory
    basename = os.path.splitext(os.path.basename(pst_file))[0]
    default_output = f"output/{basename}_extracted"

    output_dir = input(f"\n📂 Output directory [{default_output}]: ").strip()
    if not output_dir:
        output_dir = default_output

    # Ask for verbose mode
    verbose_choice = input("\n🔍 Enable verbose output? (y/n) [n]: ").strip().lower()
    verbose = verbose_choice in ["y", "yes"]

    return {
        "input": pst_file,
        "output": output_dir,
        "format": output_format,
        "verbose": verbose,
        "dry_run": False,
    }


def process_emails(args):
    """
    Process emails with the given arguments and extract them to the specified formats.

    Args:
        args (dict): A dictionary containing processing options with keys:
                    - 'input': Path to the PST/OST file to process
                    - 'output': Output directory for extracted emails
                    - 'format': Output format ('eml', 'pdf', or 'both')
                    - 'verbose': Boolean for verbose output
                    - 'dry_run': Boolean for dry run mode (optional)

    Returns:
        bool: True if processing was successful, False if errors occurred.

    This function performs the core email extraction workflow:
    - Validates the input PST/OST file
    - Creates the output directory structure
    - Loads and processes all emails from the PST file
    - Saves emails in the requested format(s)
    - Extracts and saves attachments
    - Provides progress feedback and error handling
    """
    # Validate input file
    if not os.path.exists(args["input"]):
        print(f"❌ Error: Input file '{args['input']}' does not exist.")
        return False

    if not args["input"].lower().endswith((".pst", ".ost")):
        print(f"⚠️  Warning: '{args['input']}' doesn't appear to be a PST/OST file.")

    # Create output directory
    output_dir = Path(args["output"])
    if not args.get("dry_run", False):
        output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n🔍 Processing PST/OST file: {args['input']}")
    print(f"📂 Output directory: {args['output']}")
    print(f"📄 Output format: {args['format']}")
    print("-" * 50)

    try:
        # Initialize processor and load emails
        processor = EmailProcessor(args["input"])
        emails = processor.load_emails()

        print(f"✅ Found {len(emails)} emails to process")

        if args.get("dry_run", False):
            print("\n📋 DRY RUN - Files that would be created:")
            for i, email_data in enumerate(emails[:10]):  # Show first 10 as preview
                email = email_data["message"]
                folder_path = email_data["folder_path"]

                # Format delivery time for filename
                date_prefix = processor._format_delivery_time(email)
                subject = email.subject or "no_subject"
                clean_subject = subject[:50].strip()

                if args["format"] in ["eml", "both"]:
                    print(f"  📧 {folder_path}/{date_prefix} - {clean_subject}.eml")
                if args["format"] in ["pdf", "both"]:
                    print(f"  📄 {folder_path}/{date_prefix} - {clean_subject}.pdf")

            if len(emails) > 10:
                print(f"  ... and {len(emails) - 10} more emails")
            return True

        # Initialize file saver
        file_saver = FileSaver(args["output"])

        # Process each email
        processed = 0
        for email_data in emails:
            email = email_data["message"]
            folder_path = email_data["folder_path"]

            try:
                # Create folder structure using FileSaver
                full_folder_path = file_saver._create_full_path(folder_path)

                # Save email in requested format(s)
                if args["format"] in ["eml", "both"]:
                    processor.save_as_eml(email, full_folder_path)
                    if args.get("verbose", False):
                        print(
                            f"✅ Saved EML: {folder_path}/{email.subject or 'No Subject'}"
                        )

                if args["format"] in ["pdf", "both"]:
                    processor.save_as_pdf(email, full_folder_path)
                    if args.get("verbose", False):
                        print(
                            f"✅ Saved PDF: {folder_path}/{email.subject or 'No Subject'}"
                        )

                # Save attachments if any
                try:
                    # Check if email has attachments safely
                    num_attachments = getattr(email, "number_of_attachments", 0)
                    if num_attachments and num_attachments > 0:
                        attachments_folder = os.path.join(full_folder_path, "attachments")
                        os.makedirs(attachments_folder, exist_ok=True)

                        for i in range(num_attachments):
                            try:
                                attachment = email.get_attachment(i)
                                file_saver.save_attachment(attachment, attachments_folder)
                                if args.get("verbose", False):
                                    print(f"📎 Saved attachment: {attachment.name}")
                            except Exception as e:
                                print(f"⚠️  Error saving attachment {i}: {e}")
                except Exception as e:
                    if args.get("verbose", False):
                        print(f"⚠️  Error checking attachments for email '{email.subject or 'No Subject'}': {e}")

                processed += 1

                # Progress indicator
                if not args.get("verbose", False) and processed % 50 == 0:
                    print(f"📧 Processed {processed}/{len(emails)} emails...")

            except Exception as e:
                print(
                    f"❌ Error processing email '{email.subject or 'No Subject'}': {e}"
                )
                continue

        print(f"\n🎉 Successfully processed {processed}/{len(emails)} emails!")
        print(f"📂 Files saved to: {args['output']}")
        return True

    except Exception as e:
        print(f"❌ Error: {e}")
        if args.get("verbose", False):
            import traceback

            traceback.print_exc()
        return False


def main():
    """
    Main entry point for the PST/OST Email Exporter application.

    Parses command line arguments and either runs in interactive mode or
    processes emails directly based on the provided parameters.

    Command line arguments:
    - --input/-i: Path to PST/OST file
    - --output/-o: Output directory (default: 'output')
    - --format/-f: Output format ('eml', 'pdf', 'both', default: 'eml')
    - --verbose/-v: Enable verbose output
    - --dry-run: Preview mode without saving files
    - --interactive: Force interactive mode

    Exit codes:
    - 0: Success
    - 1: Error or user cancellation
    """
    parser = argparse.ArgumentParser(
        description="Extract and convert emails from PST/OST files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  ./pst-exporter.py                                                    # Interactive mode
  ./pst-exporter.py --interactive                                      # Interactive mode
  ./pst-exporter.py -i pst_files/archive.pst                           # Command line mode
  ./pst-exporter.py --input pst_files/archive.pst --format pdf -v      # PDF format with verbose
        """,
    )

    parser.add_argument("--input", "-i", help="Path to the PST/OST file to process")

    parser.add_argument(
        "--output",
        "-o",
        default="output",
        help="Output directory for extracted emails (default: output)",
    )

    parser.add_argument(
        "--format",
        "-f",
        choices=["eml", "pdf", "both"],
        default="eml",
        help="Output format for emails (default: eml)",
    )

    parser.add_argument(
        "--verbose", "-v", action="store_true", help="Enable verbose output"
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be processed without actually saving files",
    )

    parser.add_argument(
        "--interactive", action="store_true", help="Run in interactive mode"
    )

    args = parser.parse_args()

    # If no input is provided or interactive flag is set, run interactive mode
    if not args.input or args.interactive:
        interactive_args = interactive_mode()
        if interactive_args is None:
            sys.exit(1)
        success = process_emails(interactive_args)
    else:
        # Command line mode
        cli_args = {
            "input": args.input,
            "output": args.output,
            "format": args.format,
            "verbose": args.verbose,
            "dry_run": args.dry_run,
        }
        success = process_emails(cli_args)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
