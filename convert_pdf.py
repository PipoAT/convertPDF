import PySimpleGUI as sg
import pdfplumber
from PyPDF2 import PdfReader
from openpyxl import load_workbook
import pandas as pd
import shutil
import tkinter as tk
import os
import re

def is_valid_page_input(page_numbers):
    """
    Check if the page input is not blank and contains valid data.
    """
    if not page_numbers.strip():
        return False

    try:
        # Check if all values are integers
        list(map(int, page_numbers.split(',')))
        return True
    except ValueError:
        return False
    
def merge_cells_with_text(excel_file, sheet_name):
    '''
    merge_cells_with_text() is a function that merges cells in an Excel file
    which include the one with any non-empty text followed by any and all cells
    that are blank to the right of that cell
    '''
    # Load the workbook and select the sheet
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    # Iterate over the rows in the sheet
    for row in ws.iter_rows():
        # Iterate over the cells in the row
        for cell in row:
            # If the cell contains any non-empty text
            if cell.value is not None and str(cell.value).strip() != '':
                # Find the next non-empty cell in the row
                next_cell_index = cell.column + 1
                while next_cell_index <= ws.max_column and (ws.cell(row=cell.row, column=next_cell_index).value is None or str(ws.cell(row=cell.row, column=next_cell_index).value).strip() == ''):
                    next_cell_index += 1

                # Merge the cells from the current cell to the next non-empty cell
                ws.merge_cells(start_row=cell.row, start_column=cell.column, end_row=cell.row, end_column=next_cell_index-1)

    # Save the workbook
    wb.save(excel_file)


def pdf_to_excel(pdf_file, page_numbers):
    '''
    pdf_to_excel() is a function that reads tables in a pdf file and converts the data into an excel file
    with desired formatting/restrictions
    '''
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # obtain the pdf file name
    pdf_file_name = os.path.splitext(os.path.basename(pdf_file))[0]


    # Open a file dialog to select the folder for output files
    folder_selected = os.path.dirname(pdf_file)

    # Obtain the path for the excel file
    excel_file = os.path.join(folder_selected, pdf_file_name[:10] + ".xlsx")

    excel_file_name = os.path.splitext(os.path.basename(excel_file))[0] + ".xlsx"

    # Ask the user to input the page numbers
    pages = list(map(int, page_numbers.split(',')))

    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        # Check if there are any tables on the specified pages
        has_tables = False
        for p in pages:
            page = pdf.pages[p-1]  # pdfplumber uses zero-based page indexing
            tables = page.extract_tables()  # extract all tables into a list of tables
            if any(cell.strip() for table in tables for row in table for cell in row):  # check if there are any non-empty tables that exist
                has_tables = True
                break

        # If there are tables, create the Excel file and write the tables to it
        if has_tables:
            with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                for i, p in enumerate(pages):
                    page = pdf.pages[p-1]  # pdfplumber uses zero-based page indexing
                    tables = page.extract_tables()  # extract all tables into a list of tables
                    if any(cell.strip() for table in tables for row in table for cell in row):  # check if there are any non-empty tables that exist
                        for j, table in enumerate(tables):
                            df = pd.DataFrame(table[1:], columns=table[0])  # convert each table into a pandas DataFrame

                            # format to exclude rows that have a reserved name or is '-'
                            # checks for valid column indexing and ignores formatting if invalid
                            try:
                                df = df[df[df.columns[1]] != 'Reserved']
                                df = df[df[df.columns[1]] != '—']
                            except IndexError:
                                continue

                            # sets the data to numeric if it reads a numeric value as a string, otherwise leave alone
                            for col in df.columns:
                                df[col] = df[col].apply(lambda x: pd.to_numeric(x, errors='ignore') if re.search(r'\d', str(x)) else x)

                            # write each DataFrame to the Excel file
                            df.to_excel(writer, sheet_name = pdf_file_name[:10] + '_Page' + str(p) + '_Table' + str(j+1), index=False)

        else:
            # if no tables exist, write the text
            pdf_to_txt(pdf_file, pdf_file_name, folder_selected, pages)

    # Copy the Excel file to the current directory
    shutil.copy(excel_file, os.getcwd())

    # call function to adjust format of resulting excel file to merge cells as needed
    merge_cells_with_text(excel_file_name, pdf_file_name[:10] + '_Page' + str(p) + '_Table' + str(j+1))


def pdf_to_txt(pdf_file, pdf_file_name, folder_selected, pages):
    '''
    pdf_to_txt(pdf_file, pages) is a function that reads the text of the pdf and converts into a .txt file
    if no table exists. pdf_file is a string param and pages is a list[int] param
    '''
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to select the output text file
    txt_file = os.path.join(folder_selected, pdf_file_name[:10] + ".txt")

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


# define the layout of the GUI
layout = [
    [sg.Text("Select PDF File:")],
    [sg.Input('', key="file"), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text("Enter the page numbers separated by commas (e.g., 1,2,3): ")],
    [sg.Input('', key="page")],
    [sg.Button("Convert"), sg.Button("Exit"), sg.Button("About")]
]

# create the window
window = sg.Window("PDF Conversion GUI", layout)

while True:
    # display and interact with window
    event, values = window.read()

    # if user closes window or clicks exit
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break

    # perform action with information
    if event == "Convert":
        # obtain the pdf file from user input
        pdf_file = values['file']
        # obtain page numbers from user input
        page_numbers = values['page']
        # checks if the page numbers are valid inputs by user
        if not is_valid_page_input(page_numbers):
            # throw popup if user input for page numbers is blank or invalid
            sg.popup("ALERT", "Please enter valid page numbers!")
        else:
            # call the function to convert if pdf is valid
            if pdf_file:
                pdf_to_excel(pdf_file, page_numbers)
                # show alert that conversion is complete
                sg.popup("ALERT","Conversion is a success!")
                # clear input fields
                window['file'].update('')
                window['page'].update('')
            else:
                # if no pdf is selected or is invalid, display alert
                sg.popup("ALERT","Please select a PDF file")
    
    # if user selects the About button, display the software info as a popup/alert
    if event == "About":
        sg.popup("INFO","Version 1.0.0\nSoftware Developed by Andrew T. Pipo")

# close window
window.close()