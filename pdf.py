import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, Listbox, Toplevel
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pytesseract
from PIL import Image
import fitz
import os
import pdf2docx
import pandas as pd
from pdf2image import convert_from_path
import tabula
import time  # For simulating long operations (remove in production)
import webbrowser 
import PyPDF2  # For handling PDF password setting and text search

# Set the Tesseract executable path for Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\HP\scoop\apps\tesseract\current\tesseract.exe'  # Update with your Tesseract path

# Function to extract text using OCR from the PDF pages (image-based PDFs)
def extract_text_ocr(file_path):
    doc = fitz.open(file_path)  # Open the PDF using PyMuPDF
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)  # Load a page
        pix = page.get_pixmap()  # Convert page to an image (for OCR)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)  # Convert to PIL Image

        # Use pytesseract to extract text from the image
        page_text = pytesseract.image_to_string(img)

        text += f"Page {page_num + 1}:\n{page_text}\n\n"  # Add page number for clarity
    return text

# Function to merge multiple PDFs
def merge_pdfs(file_paths, output_path, progress_var):
    merger = fitz.open()  # Using fitz for handling PDFs
    total_files = len(file_paths)
    for index, pdf in enumerate(file_paths):
        merger.insert_pdf(fitz.open(pdf))
        progress_var.set((index + 1) / total_files * 100)  # Update progress
        time.sleep(0.5)  # Simulate processing time (remove in production)
    merger.save(output_path)
    merger.close()

# Function to convert PDF to Word
def pdf_to_word(pdf_file, output_path, progress_var):
    try:
        pdf2docx.PdfToDocx().convert(pdf_file, output_path)
        progress_var.set(100)  # Set progress to 100% after completion
    except Exception as e:
        print(f"An error occurred while converting to Word: {str(e)}")

# Function to convert PDF to Excel using tabula-py
def pdf_to_excel(pdf_file, output_path, progress_var):
    try:
        tables = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
        with pd.ExcelWriter(output_path) as writer:
            for i, table in enumerate(tables):
                table.to_excel(writer, sheet_name=f"Sheet{i + 1}", index=False)
                progress_var.set((i + 1) / len(tables) * 100)  # Update progress
        progress_var.set(100)  # Set progress to 100% after completion
    except Exception as e:
        print(f"An error occurred while converting to Excel: {str(e)}")

# Function to set a password for a PDF
def set_pdf_password(input_pdf_path, output_pdf_path, password):
    try:
        with open(input_pdf_path, 'rb') as input_pdf_file:
            reader = PyPDF2.PdfReader(input_pdf_file)
            writer = PyPDF2.PdfWriter()

            # Add all pages to the writer
            for page in reader.pages:
                writer.add_page(page)

            # Set the password on the output file
            writer.encrypt(password)

            # Write the encrypted PDF to the output file
            with open(output_pdf_path, 'wb') as output_pdf_file:
                writer.write(output_pdf_file)

        return True
    except Exception as e:
        print(f"Error setting password: {e}")
        return False

# Function to search for text in a PDF
def search_text_in_pdf(pdf_path, search_text):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            found_text = []

            # Search through each page of the PDF
            for page_num, page in enumerate(reader.pages):
                page_text = page.extract_text()
                if page_text and search_text.lower() in page_text.lower():
                    found_text.append(f"Found on page {page_num + 1}")

            return "\n".join(found_text) if found_text else "Text not found."

    except Exception as e:
        print(f"Error searching text: {e}")
        return None

# Function to convert PDF to images using PyMuPDF (fitz)
def pdf_to_images_with_fitz(pdf_file, output_folder):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_file)

        # Iterate through each page of the PDF
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)

            # Render page as an image (pixmap)
            pix = page.get_pixmap()

            # Save the image
            output_image_path = f"{output_folder}/page_{page_num + 1}.png"
            pix.save(output_image_path)

        return True
    except Exception as e:
        print(f"Error converting PDF to images: {e}")
        return False

# GUI for PDF Utility
class PDFUtilityGUI:
    def __init__(self, master):
        self.master = master
        master.title("PDF Utility")
        master.geometry("900x600")
        master.resizable(True, True)

        # Create a style
        self.style = ttk.Style(theme="darkly")

        # Main frame
        self.main_frame = ttk.Frame(master, padding="20 20 20 20")
        self.main_frame.pack(fill=BOTH, expand=YES)

        # Title
        ttk.Label(self.main_frame, text="PDF Utility", font=("Helvetica", 24, "bold")).pack(pady=20)

        # Operation selection frame
        self.operation_frame = ttk.LabelFrame(self.main_frame, text="Select Operation", padding="20 20 20 20")
        self.operation_frame.pack(fill=X, padx=10, pady=10)

        self.operation = tk.StringVar(value="search")
        operations = [
            ("Extract Text", "extract"),
            ("Merge PDFs", "merge"),
            ("Convert to Images", "convert_images"),
            ("Convert to Word", "convert_word"),
            ("Convert to Excel", "convert_excel"),
            ("Set Password", "set_password"),
            ("Search Text", "search_text"),
        ]

        for text, value in operations:
            ttk.Radiobutton(self.operation_frame, text=text, variable=self.operation, value=value).pack(side=LEFT, padx=10)

        # File selection frame
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=X, padx=10, pady=20)

        self.file_button = ttk.Button(self.file_frame, text="Select PDF File(s)", command=self.browse_files, width=30)
        self.file_button.pack(side=LEFT, padx=30)

        self.file_label = ttk.Label(self.file_frame, text="No files selected")
        self.file_label.pack(side=LEFT, padx=10)

        # Process button
        self.process_button = ttk.Button(self.main_frame, text="Process", command=self.process, style="success.TButton", width=40)
        self.process_button.pack(pady=40)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=X, padx=10, pady=10)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        self.status_bar.pack(fill=X, side=BOTTOM, pady=15)

        self.selected_files = []


    # Footer
        self.footer_frame = ttk.Frame(master)
        self.footer_frame.pack(side=BOTTOM, fill=X, pady=(10, 0))

        footer_label = ttk.Label(self.footer_frame, text="Created By : ", font=("Helvetica", 12))
        footer_label.pack(side=LEFT, padx=5)

        link_label = ttk.Label(self.footer_frame, text="Aditya Nalawade", font=("Helvetica", 12, "underline"), foreground="#CBFC01")
        link_label.pack(side=LEFT)

        link_label2 = ttk.Label(self.footer_frame, text="GitHub", font=("Helvetica", 12, "underline"), foreground="#CBFC01")
        link_label2.pack(side=LEFT, padx=5)
        link_label2.bind("<Button-1>", self.open_link2)

        link_label3 = ttk.Label(self.footer_frame, text="Mail", font=("Helvetica", 12, "underline"), foreground="#CBFC01")
        link_label3.pack(side=LEFT, padx=5)
        link_label3.bind("<Button-1>", self.open_link3)

        link_label.bind("<Button-1>", self.open_link)

    def open_link(self, event):
        webbrowser.open("https://www.linkedin.com/in/aditya-nalawade-a4b081297?utm_source=share&utm_medium=member_desktop")

    def open_link2(self, event):
        webbrowser.open("https://github.com/aditya-nalawade")

    def open_link3(self, event):
        webbrowser.open("mailto:nalawadeaditya01@gmail.com")


    def browse_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            self.selected_files = file_paths
            self.file_label.config(text=f"{len(file_paths)} file(s) selected")

    def process(self):
        operation = self.operation.get()
        if not self.selected_files:
            messagebox.showerror("Error", "Please select a PDF file")
            return

        # Initialize variables
        progress_var = self.progress_var
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        output_folder = filedialog.askdirectory()

        # Process based on operation
        try:
            if operation == "extract":
                output_text = extract_text_ocr(self.selected_files[0])
                self.status_var.set("Text extracted successfully!")
                with open(output_path, "w") as output_file:
                    output_file.write(output_text)

            elif operation == "merge":
                merge_pdfs(self.selected_files, output_path, progress_var)
                self.status_var.set(f"PDFs merged successfully!")

            elif operation == "convert_images":
                success = pdf_to_images_with_fitz(self.selected_files[0], output_folder)
                if success:
                    self.status_var.set(f"Converted PDF to images successfully!")

            elif operation == "convert_word":
                pdf_to_word(self.selected_files[0], output_path, progress_var)
                self.status_var.set(f"Converted PDF to Word successfully!")

            elif operation == "convert_excel":
                pdf_to_excel(self.selected_files[0], output_path, progress_var)
                self.status_var.set(f"Converted PDF to Excel successfully!")

            elif operation == "set_password":
                password = simpledialog.askstring("Password", "Enter password:")
                if set_pdf_password(self.selected_files[0], output_path, password):
                    self.status_var.set("Password set successfully!")

            elif operation == "search_text":
                search_text = simpledialog.askstring("Search", "Enter text to search:")
                result = search_text_in_pdf(self.selected_files[0], search_text)
                if result:
                    messagebox.showinfo("Search Result", result)
                    self.status_var.set("Text search completed!")

        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")


# Run the application
root = ttk.Window(themename="darkly")
app = PDFUtilityGUI(root)
root.mainloop()
