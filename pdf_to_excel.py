import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

'''
pip install pdfplumber
    
pip install pandas

pip install openpyxl

pip install xlsxwriter

'''

def pdf_to_excel():
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to select the PDF file
    pdf_file = filedialog.askopenfilename(title="Select the PDF file", filetypes=[("PDF Files", "*.pdf")])
    # Open a file dialog to select the output Excel file
    excel_file = filedialog.asksaveasfilename(title="Select or name and save a new Excel file", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

    # Obtain the pdf name
    pdf_file_name = os.path.splitext(os.path.basename(pdf_file))[0]

    # Ask the user to input the page numbers
    pages = input("Enter the page numbers separated by commas (e.g., 1,2,3): ")
    pages = list(map(int, pages.split(',')))

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        # Create a Pandas Excel writer using XlsxWriter as the engine
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            for i, p in enumerate(pages):
                page = pdf.pages[p-1]  # pdfplumber uses zero-based page indexing
                table = page.extract_table()  # extract the table into a list of lists
                df = pd.DataFrame(table[1:], columns=table[0])  # convert the table into a pandas DataFrame
                df = df[df['Name'] != 'Reserved']
                df.to_excel(writer, sheet_name = pdf_file_name[:29] + '_' + str(i), index=False)  # write the DataFrame to the Excel file

# Use the function
pdf_to_excel()
