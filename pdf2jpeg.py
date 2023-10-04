from pdf2image import convert_from_path
from tkinter import filedialog


def select_output_folder():
    output_folder_path = filedialog.askdirectory()
    return output_folder_path


def pdf_to_jpeg(pdf_file, output):
    output_folder = output
    if output_folder:
        pages = convert_from_path(pdf_file, dpi=300)  # Adjust dpi as needed
        for i, page in enumerate(pages):
            page.save(f"{output_folder}/page_{i + 1}.jpeg", 'JPEG')
        return 1
    else:
        return 0