#!/usr/bin/env python3
"""
PST/OST Email Exporter - Main Script
This script processes PST/OST files and extracts emails to various formats.
"""

import argparse
import os
import sys

# Add src directory to path to import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from email_processor import EmailProcessor  # noqa: E402
from pst_processor import PSTProcessor  # noqa: E402


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
  ./pst-exporter.py                                   # Interactive mode
  ./pst-exporter.py --interactive                     # Interactive mode
  ./pst-exporter.py -i pst_files/archive.pst          # Command line mode
  ./pst-exporter.py --input pst_files/archive.pst \\  # PDF format with
    --format pdf -v                                   # verbose output
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
        interactive_args = _interactive_mode()
        if interactive_args is None:
            sys.exit(1)
        success = _process_emails(interactive_args)
    else:
        # Command line mode
        cli_args = {
            "input": args.input,
            "output": args.output,
            "format": args.format,
            "verbose": args.verbose,
            "dry_run": args.dry_run,
        }
        success = _process_emails(cli_args)

    sys.exit(0 if success else 1)


def _interactive_mode():
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

    # Find PST files using PSTProcessor
    pst_processor = PSTProcessor()
    pst_files = pst_processor.find_pst_files()

    if not pst_files:
        print("⚠️  No PST/OST files found in 'pst_files' directory")
        print(
            "Please place your PST/OST files in the 'pst_files' directory "
            "and run again."
        )

        # Ask if user wants to specify custom path
        custom = (
            input("\nWould you like to specify a custom path? (y/n): ").lower().strip()
        )
        if custom in ["y", "yes"]:
            pst_file = input("Enter the full path to your PST/OST file: ").strip()
            if not pst_processor.validate_pst_file(pst_file):
                print(
                    f"❌ File '{pst_file}' does not exist or is not a valid PST/OST file"
                )
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
                if not pst_processor.validate_pst_file(pst_file):
                    print(
                        f"❌ File '{pst_file}' does not exist or is not a "
                        f"valid PST/OST file"
                    )
                    continue
                break

            try:
                file_num = int(choice)
                if 1 <= file_num <= len(pst_files):
                    pst_file = pst_files[file_num - 1]
                    break
                print(
                    f"❌ Invalid choice. Please enter a number between 1 "
                    f"and {len(pst_files)}"
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
        if format_choice == "2":
            output_format = "pdf"
            break
        if format_choice == "3":
            output_format = "both"
            break
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


def _process_emails(args):
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

    This function is a wrapper that delegates to EmailProcessor.process_emails().
    """
    # Initialize processor and delegate to the EmailProcessor method
    processor = EmailProcessor(args["input"])
    return processor.process_emails(
        output_dir=args["output"],
        output_format=args["format"],
        verbose=args.get("verbose", False),
        dry_run=args.get("dry_run", False),
    )


if __name__ == "__main__":
    main()
