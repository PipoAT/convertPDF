from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog


# Required dependencies:

## Uses Python 3.6 or newer


## tkinter included with Python install


# INSTALL VIA PIP:
'''
pip install PyPDF2
    
pip install pycryptodome

'''



def pdf_to_txt():
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to select the PDF file
    pdf_file = filedialog.askopenfilename(title="Select the PDF file", filetypes=[("PDF Files", "*.pdf")])
    # Open a file dialog to select the output text file
    txt_file = filedialog.asksaveasfilename(title="Name and save the output text file", defaultextension=".txt", filetypes=[("Text Files", "*.txt")])

    # Ask the user for the start and end page numbers
    start_page = int(input("Enter the start page number: ")) - 1  # Subtract 1 because page numbers start from 0
    end_page = int(input("Enter the end page number: "))

    # Create a PDF file reader object
    pdf_reader = PdfReader(pdf_file)

    # Initialize an empty string to hold the PDF text
    pdf_text = ''

    # Loop through the selected pages in the PDF
    for page in pdf_reader.pages[start_page:end_page]:
        # Extract the text from the page
        page_text = page.extract_text()
        # Add the page text to the overall PDF text
        pdf_text += page_text

    # Open the text file in write mode with utf-8 encoding
    with open(txt_file, 'w', encoding='utf-8') as txt:
        # Write the PDF text to the text file
        txt.write(pdf_text)

# Use the function
pdf_to_txt()
