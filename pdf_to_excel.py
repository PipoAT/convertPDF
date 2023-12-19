import pdfplumber
from PyPDF2 import PdfReader
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re

'''
pip install pdfplumber
    
pip install pandas

pip install openpyxl

pip install xlsxwriter

pip install PyPDF2

pip install pycryptodome

'''


'''
pdf_to_excel() is a function that reads tables in a pdf file and converts the data into an excel file
with desired formatting/restrictions
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
                tables = page.extract_tables()  # extract all tables into a list of tables
                if any(cell.strip() for table in tables for row in table for cell in row):  # check if there are any non-empty tables that exist
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table[1:], columns=table[0])  # convert each table into a pandas DataFrame

                        # format to exclude rows that have a reserved name or is '-'
                        df = df[df[df.columns[1]] != 'Reserved']
                        df = df[df[df.columns[1]] != '—']

                        # sets the data to numeric if it reads a numeric value as a string, otherwise leave alone
                        for col in df.columns:
                            df[col] = df[col].apply(lambda x: pd.to_numeric(x, errors='ignore') if re.search(r'\d', str(x)) else x)

                        # write each DataFrame to the Excel file
                        df.to_excel(writer, sheet_name = pdf_file_name[:10] + '_Page' + str(p) + "_" + str(i+1) + '_Table' + str(j+1), index=False)
                else:
                    break
        # if no tables exist, write the text
        pdf_to_txt(pdf_file, pages)



'''
pdf_to_txt(pdf_file, pages) is a function that reads the text of the pdf and converts into a .txt file
if no table exists. pdf_file is a string param and pages is a list[int] param
'''
def pdf_to_txt(pdf_file, pages):
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to select the output text file
    txt_file = filedialog.asksaveasfilename(title="Name and save the output text file", defaultextension=".txt", filetypes=[("Text Files", "*.txt")])

    # Ask the user for the start and end page numbers
    start_page = min(pages)-1
    end_page = max(pages)

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
pdf_to_excel()
