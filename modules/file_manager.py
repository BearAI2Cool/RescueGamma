import os
import shutil
import tempfile
import zipfile


class FileManager:
    def __init__(self):
        self.temp_dir = None

    def extract_pptx(self, pptx_path: str) -> str:
        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"PPT文件不存在: {pptx_path}")

        self.temp_dir = tempfile.mkdtemp()

        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(self.temp_dir)

        return self.temp_dir

    def compress_to_pptx(self, temp_dir: str, output_path: str):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)

    def cleanup(self):
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
            self.temp_dir = None

    def get_slide_files(self, temp_dir: str) -> list:
        slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
        if not os.path.exists(slides_dir):
            return []

        slide_files = []
        for file in os.listdir(slides_dir):
            if file.startswith('slide') and file.endswith('.xml'):
                slide_files.append(os.path.join(slides_dir, file))

        return sorted(slide_files)
