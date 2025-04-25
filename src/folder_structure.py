class FolderStructure:
    def __init__(self, base_path):
        self.base_path = base_path

    def create_folder(self, folder_name):
        import os
        path = os.path.join(self.base_path, folder_name)
        os.makedirs(path, exist_ok=True)
        return path

    def maintain_structure(self, folder_hierarchy):
        for folder in folder_hierarchy:
            self.create_folder(folder)