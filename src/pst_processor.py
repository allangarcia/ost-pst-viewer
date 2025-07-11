import glob
import os


class PSTProcessor:
    """
    PSTProcessor handles PST/OST file discovery and management operations.

    This class provides methods to:
    - Find PST/OST files in specified directories
    - Validate PST/OST file paths
    - Manage PST file discovery directories

    Attributes:
        pst_directory (str): Directory where PST/OST files are located.
    """

    def __init__(self, pst_directory="pst_files"):
        """
        Initialize the PSTProcessor with a PST files directory.

        Args:
            pst_directory (str, optional): Directory containing PST/OST files.
                Defaults to "pst_files".
        """
        self.pst_directory = pst_directory

    def find_pst_files(self):
        """
        Find all PST/OST files in the configured PST directory.

        Returns:
            list: A sorted list of file paths to PST/OST files found in the
                PST directory. Returns an empty list if no files are found
                or if the directory doesn't exist.

        Note:
            Creates the PST directory if it doesn't exist.
        """
        if not os.path.exists(self.pst_directory):
            os.makedirs(self.pst_directory, exist_ok=True)
            return []

        pst_files = glob.glob(os.path.join(self.pst_directory, "*.pst")) + glob.glob(
            os.path.join(self.pst_directory, "*.ost")
        )
        return sorted(pst_files)

    def validate_pst_file(self, file_path):
        """
        Validate that a given file path points to a valid PST/OST file.

        Args:
            file_path (str): Path to the PST/OST file to validate.

        Returns:
            bool: True if the file exists and has a valid PST/OST extension,
                False otherwise.
        """
        if not file_path or not os.path.exists(file_path):
            return False

        file_extension = os.path.splitext(file_path)[1].lower()
        return file_extension in [".pst", ".ost"]

    def get_pst_files_count(self):
        """
        Get the count of PST/OST files in the configured directory.

        Returns:
            int: Number of PST/OST files found in the directory.
        """
        return len(self.find_pst_files())

    def ensure_pst_directory_exists(self):
        """
        Ensure the PST directory exists, creating it if necessary.

        Returns:
            bool: True if directory exists or was created successfully,
                False if creation failed.
        """
        try:
            if not os.path.exists(self.pst_directory):
                os.makedirs(self.pst_directory, exist_ok=True)
            return True
        except OSError as e:
            print(f"Error creating PST directory '{self.pst_directory}': {e}")
            return False
