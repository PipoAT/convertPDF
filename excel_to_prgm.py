import openpyxl
import tkinter as tk
from tkinter import filedialog

# Create a root window and hide it
root = tk.Tk()
root.withdraw()

# Open a file dialog and get the selected file path
file_path = filedialog.askopenfilename()

# Load the workbook and select the active worksheet
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Open a text file in write mode
with open('output.txt', 'w') as file:
    # Get the headers from the first row of the worksheet
    headers = [cell.value for cell in sheet[1]]

    # Check if 'address' and 'page' are in the headers and get their indices
    indices_to_exclude = [i for i, header in enumerate(headers) if header.lower() in ['address', 'page']]

    # Iterate through each row in the worksheet starting from the second row
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Skip rows without a name
        if row[0] is None:
            continue
        # Write #define directives to text file
        j = 8
        for i, cell in enumerate(row[2:10], 1):
            # Skip the cells in the 'address' and 'page' columns
            if i in indices_to_exclude or cell is None or cell == '-':
                continue
            if ' ' in cell:
                continue
            file.write(f'#define {row[1]}_Bit{j-1} {cell}\n')
            j -= 1
