import PySimpleGUI as sg
import pdfplumber
from openpyxl import load_workbook
import pandas as pd
import os
import re

def excel_to_prgm(wbactivesheet):
    """
    Converts the formatted excel spreadsheet into the cpp file as #define headers with desired formatting
    """
    # Open a text file in write mode
    with open('output.cpp', 'w') as file:
        # Get the headers from the first row of the worksheet
        headers = [cell.value for cell in wbactivesheet[1]]

        # Check if 'address' and 'page' are in the headers and get their indices
        indices_to_exclude = [i for i, header in enumerate(headers) if 'address' in header.lower() or 'page' in header.lower()]

        # Iterate through each row in the worksheet starting from the second row
        for row in wbactivesheet.iter_rows(min_row=2, values_only=True):
            # Skip rows without a name
            if row[0] is None:
                continue
            # Write #define directives to cpp file
            for i, cell in enumerate(row[2:10], 1):
                # Skip the cells with invalid data
                if i in indices_to_exclude or not cell or cell in ['-', '—', ' ']:
                    continue
                row_as_list = list(row)  # Convert tuple to list
                if '(' in row_as_list[1]: 
                    # remove any () as they relate to footnotes in datasheets and are not needed in cpp file
                    row_as_list[1] = re.sub(r'\(.*?\)', '', row_as_list[1])
                row = tuple(row_as_list)  # Convert list back to tuple if necessary
                if '(' in cell:
                    cell = re.sub(r'\(.*?\)', '', cell)
                if 'PIN' in cell:
                    file.write(f'#define {cell} {cell[3]}, {cell[4]}\n')
                elif 'PORT' in cell:
                    file.write(f'#define {cell} {cell[4]}, {cell[5]}\n')
                elif 'DD' in cell:
                    file.write(f'#define {cell} {cell[2]}, {cell[3]}\n')
                else:
                    file.write(f'#define {row[1]}_Bit{8-i} {cell}\n')


def is_valid_page_input(page_numbers):
    """
    Check if the page input is not blank and contains valid data.
    """
    if not page_numbers.strip():
        return False
    try:
        # Check if all values are integers
        list(map(int, page_numbers))
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
    wbactivesheet = wb.active
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
    # convert to .cpp file with define statements
    excel_to_prgm(wbactivesheet)


def pdf_to_excel(pdf_file, page_numbers):
    '''
    pdf_to_excel() is a function that reads tables in a pdf file and converts the data into an excel file
    with desired formatting/restrictions
    '''

    # obtain the pdf file name
    pdf_file_name = os.path.splitext(os.path.basename(pdf_file))[0]

    # Open a file dialog to select the folder for output files
    folder_selected = os.path.dirname(pdf_file)

    # Obtain the path for the excel file
    excel_file = os.path.join(folder_selected, pdf_file_name[:10] + ".xlsx")

    # Ask the user to input the page numbers
    pages = list(map(int, page_numbers.split(',')))

    has_tables = False
    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        # If there are tables, create the Excel file and write the tables to it
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            for _, p in enumerate(pages):
                page = pdf.pages[p-1]  # pdfplumber uses zero-based page indexing
                tables = page.extract_tables()  # extract all tables into a list of tables
                if any(cell.strip() for table in tables for row in table for cell in row):  # check if there are any non-empty tables that exist
                    has_tables = True
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table[1:], columns=table[0])  # convert each table into a pandas DataFrame

                        # format to exclude rows that have a reserved name or is '-' assuming valid column exists
                        try:
                            df = df[(df[df.columns[1]] != 'Reserved') & (df[df.columns[1]] != '—')]
                        except IndexError:
                            continue

                        # sets the data to numeric if it reads a numeric value as a string, otherwise leave alone
                        for col in df.columns:
                            df[col] = df[col].apply(lambda x: pd.to_numeric(x, errors='ignore') if re.search(r'\d', str(x)) else x)

                        # write each DataFrame to the Excel file
                        df.to_excel(writer, sheet_name = pdf_file_name[:10] + '_Page' + str(p) + '_Table' + str(j+1), index=False)
                else:
                    # if no tables exist, write the text
                    sg.popup("ALERT", "No Tables Exist.")
                    break

    # call function to adjust format of resulting excel file to merge cells as needed
    if has_tables:
        merge_cells_with_text(excel_file, pdf_file_name[:10] + '_Page' + str(p) + '_Table' + str(j+1))
    else:
        os.remove(excel_file)

# define the layout of the GUI
layout = [
    [sg.Text("Select PDF File:")],
    [sg.Input('', key="file"), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text("Enter the page number: ")],
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
                window['page'].update('')
            else:
                # if no pdf is selected or is invalid, display alert
                sg.popup("ALERT","Please select a PDF file")
    
    # if user selects the About button, display the software info as a popup/alert
    if event == "About":
        sg.popup("INFO","Version 1.0.0\nSoftware Developed by Andrew T. Pipo")

# close window
window.close()