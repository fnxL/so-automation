import os


def get_pdf_files_in_path(path: str):
    return [f for f in os.listdir(path) if f.lower().endswith(".pdf")]
