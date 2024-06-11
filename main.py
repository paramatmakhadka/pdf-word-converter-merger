import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pdf2docx import Converter
from docx2pdf import convert
from PyPDF2 import PdfMerger

merge_count = 1

# Function to convert PDF to Word
def pdf_to_word(pdf_path):
    save_dir = os.path.dirname(pdf_path)
    base_name = os.path.basename(pdf_path)
    word_name = os.path.splitext(base_name)[0] + ".docx"
    word_path = os.path.join(save_dir, word_name)
    cv = Converter(pdf_path)
    cv.convert(word_path)
    cv.close()
    messagebox.showinfo("Success", f"Converted {pdf_path} to {word_path}")

# Function to convert Word to PDF
def word_to_pdf(word_path):
    save_dir = os.path.dirname(word_path)
    base_name = os.path.basename(word_path)
    pdf_name = os.path.splitext(base_name)[0] + ".pdf"
    pdf_path = os.path.join(save_dir, pdf_name)
    convert(word_path, pdf_path)
    messagebox.showinfo("Success", f"Converted {word_path} to {pdf_path}")

# Function to merge PDFs
def merge_pdfs(pdf_paths):
    global merge_count
    if not pdf_paths:
        return
    save_dir = os.path.dirname(pdf_paths[0])
    merger = PdfMerger()
    for pdf in pdf_paths:
        merger.append(pdf)
    merged_pdf_name = f"merged_pdf{merge_count}.pdf"
    merged_pdf_path = os.path.join(save_dir, merged_pdf_name)
    merger.write(merged_pdf_path)
    merger.close()
    merge_count += 1
    messagebox.showinfo("Success", f"Merged PDFs into {merged_pdf_path}")

# Function to handle PDF to Word conversion
def handle_pdf_to_word():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        pdf_to_word(pdf_path)

# Function to handle Word to PDF conversion
def handle_word_to_pdf():
    word_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if word_path:
        word_to_pdf(word_path)

# Function to handle merging PDFs
def handle_merge_pdfs():
    pdf_paths = []
    while True:
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            break
        pdf_paths.append(pdf_path)
        more_files = messagebox.askyesno("Add More PDFs", "Do you want to add another PDF?")
        if not more_files:
            break
    if pdf_paths:
        merge_pdfs(pdf_paths)

# Creating the main window
root = tk.Tk()
root.title("PDF and Word Converter")

# Setting the style
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12), padding=10)
style.configure("TLabel", font=("Helvetica", 14), padding=10)
root.configure(bg="#f0f0f0")

# Creating a frame for the buttons
frame = ttk.Frame(root, padding="20 20 20 20", style="TFrame")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Creating buttons for each function
btn_pdf_to_word = ttk.Button(frame, text="Convert PDF to Word", command=handle_pdf_to_word)
btn_word_to_pdf = ttk.Button(frame, text="Convert Word to PDF", command=handle_word_to_pdf)
btn_merge_pdfs = ttk.Button(frame, text="Merge PDFs", command=handle_merge_pdfs)

# Placing the buttons on the frame
btn_pdf_to_word.grid(row=1, column=0, pady=10, sticky=(tk.W, tk.E))
btn_word_to_pdf.grid(row=2, column=0, pady=10, sticky=(tk.W, tk.E))
btn_merge_pdfs.grid(row=3, column=0, pady=10, sticky=(tk.W, tk.E))

# Adding a label for the creator's name
creator_label = ttk.Label(root, text="Created by Paramatma Khadka", font=("Helvetica", 10), background="#f0f0f0")
creator_label.grid(row=1, column=0, pady=(10, 20), sticky=(tk.W, tk.E))

# Running the GUI main loop
root.mainloop()
