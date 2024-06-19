import os
from pdf2docx import parse
from docx2pdf import convert

def convert_file(file_path):

    filename, file_extension = os.path.splitext(file_path)

    if file_extension.lower() in (".docx", ".doc"): # Convert Word to PDF
        output_file_path = f"{filename}.pdf"
        try:
            convert(file_path, output_file_path)
            print(f"Successfully converted Word file to PDF: {output_file_path}")
            open_file_explorer(output_file_path)  # Open File Explorer at the PDF's location
        except Exception as e:
            print(f"Error converting Word to PDF: {e}")

    elif file_extension.lower() == ".pdf":  # Convert PDF to Word using pdf2docx
        output_file_path = f"{filename}.docx"
        try:
            parse(pdf_file=file_path, docx_with_path=output_file_path)
            print(f"Successfully converted PDF to Word: {output_file_path}")
            open_file_explorer(output_file_path)  # Open File Explorer at the Word's location
        except Exception as e:
            print(f"Error converting PDF to Word: {e}")

    else:
        print(f"Unsupported file type: {file_extension}")

def open_file_explorer(file_path):

    directory = os.path.dirname(file_path)  # Get the directory containing the file
    if os.name == 'nt':  # Windows
        os.startfile(directory)
    elif os.name == 'posix':  # Linux, macOS
        os.system(f'xdg-open {directory}')  # Open the directory containing the file

# Get the file path using a file dialog (hidden)
try:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename()
    root.destroy()  # Destroy the hidden window

    if file_path:
        convert_file(file_path)

except ImportError:
    print("Please run this script in an environment with tkinter installed.")


