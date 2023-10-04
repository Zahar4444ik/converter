import os
import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert
from pdf2docx import parse
from docx import Document
import PyPDF2
from pdf2image import convert_from_path


def update_destination_formats(*args):  # Показує варінати другого меню в залежності від першого
    selected_source_format = source_format_var.get()

    format_mappings = {
        "PDF": ["DOCX", "JPEG", "Text"],
        "DOCX": ["PDF", "Text"],
        "JPEG": ["PDF"],
        "Text": ["PDF", "DOCX"]
    }

    # Get the currently selected destination format
    selected_destination_format = destination_format_var.get()

    # Clear existing options in the destination format dropdown
    destination_format_menu['menu'].delete(0, 'end')

    # Add the updated options based on the selected source format
    for dest_format in format_mappings.get(selected_source_format, []):
        destination_format_menu['menu'].add_command(label=dest_format,
                                                    command=tk._setit(destination_format_var, dest_format))

    # If the previously selected destination format is still valid, set it
    if selected_destination_format in format_mappings.get(selected_source_format, []):
        destination_format_var.set(selected_destination_format)
    else:
        # Otherwise, set the default destination format
        destination_format_var.set("")


def converter():
    selected_filetype = source_format_var.get()
    selected_filetype2 = destination_format_var.get()
    filetypes = [("DOCX files", "*.docx"), ("All files", "*.*")]
    if selected_filetype == "PDF":
        filetypes = [("PDF files", "*.pdf"), ("All files", "*.*")]
    elif selected_filetype == "DOCX":
        filetypes = [("DOCX files", "*.docx"), ("All files", "*.*")]
    if selected_filetype == ' ' or selected_filetype2 == ' ':  # Перевірка чи варінати вибрані
        source_format_var.set(" ")
        destination_format_var.set(" ")
        label.config(text='Choose files')
        return 0
    file = filedialog.askopenfilename(filetypes=filetypes)
    if file:
        if selected_filetype == "DOCX":  # Convert DOCX files
            if selected_filetype2 == 'PDF':
                pdf_file = os.path.splitext(file)[0] + ".pdf"
                convert(file, pdf_file)
                label.config(text=f"Converted {file} \nto {pdf_file}")
                source_format_var.set(" ")
                destination_format_var.set(" ")
            else:
                doc = Document(file)
                text = []
                for paragraph in doc.paragraphs:
                    text.append(paragraph.text)
                path = os.path.splitext(file)[0] + ".txt"
                with open(path, 'w') as text_file:
                    text_file.write(''.join(text))
                label.config(text=f"Converted {file} \nto {path}")
                source_format_var.set(" ")
                destination_format_var.set(" ")

        elif selected_filetype == "PDF":  # Convert PDF files
            if selected_filetype2 == 'DOCX':
                docx_file = os.path.splitext(file)[0] + ".docx"
                parse(file, docx_file)
                label.config(text=f"Converted {file} \nto {docx_file}")
                source_format_var.set(" ")
                destination_format_var.set(" ")
            elif selected_filetype2 == 'Text':
                text = ''
                with open(file, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for page_number in range(pdf_reader.getNumPages()):
                            page = pdf_reader.getPage(page_number)
                            text += page.extractText()
                path = os.path.splitext(file)[0] + ".txt"
                with open(path, 'w') as txt:
                    txt.write(text)
                label.config(text=f"Converted {file} \nto {path}")
                source_format_var.set(" ")
                destination_format_var.set(" ")
            else:
                pdf_dir = os.path.dirname(file)

                pages = convert_from_path(file)
                for i, page in enumerate(pages):
                    page.save(os.path.join(pdf_dir, f"page_{i + 1}.jpeg"), 'JPEG')


# Create the main window
root = tk.Tk()
root.title("File Format Converter")

# Create a label for source format
source_format_label = tk.Label(root, text="Convert from:")
source_format_label.pack()

# Create a variable to store the selected source format
source_format_var = tk.StringVar()
source_format_var.set(" ")  # Set the initial source format

# Create the source format dropdown menu
source_format_menu = tk.OptionMenu(root, source_format_var, "PDF", "DOCX", "JPEG", "Text")
source_format_menu.pack()

# Create a label for destination format
destination_format_label = tk.Label(root, text="Convert to:")
destination_format_label.pack()

# Create a variable to store the selected destination format
destination_format_var = tk.StringVar()
destination_format_var.set(' ')
destination_format_menu = tk.OptionMenu(root, destination_format_var, "")
destination_format_menu.pack()

# Create a button for conversion (you can add your conversion logic here)
convert_button = tk.Button(root, text="Convert", command=converter)
convert_button.pack()
label = tk.Label(root, text='')
label.pack()

# Set a trace on the source format variable to update destination formats
source_format_var.trace('w', update_destination_formats)

# Start the Tkinter main loop
root.mainloop()
