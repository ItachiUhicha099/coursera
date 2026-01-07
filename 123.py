import os
import datetime
import openpyxl
import win32com.client as win32

# === CONFIG ===
folder_path = r"C:\Users\YourName\Documents\excel_folder"
file_name = "Report.xlsx"
file_path = os.path.join(folder_path, file_name)

sheet_name = "Sheet1"
cell_range = "A1:C5"   # change as needed

# === STEP 1: Load Excel and read range ===
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb[sheet_name]

# Parse range
start_cell, end_cell = cell_range.split(":")
start_col = openpyxl.utils.column_index_from_string(''.join([c for c in start_cell if c.isalpha()]))
start_row = int(''.join([c for c in start_cell if c.isdigit()]))
end_col = openpyxl.utils.column_index_from_string(''.join([c for c in end_cell if c.isalpha()]))
end_row = int(''.join([c for c in end_cell if c.isdigit()]))

# Extract values
data = []
for row in ws.iter_rows(min_row=start_row, max_row=end_row,
                        min_col=start_col, max_col=end_col, values_only=True):
    data.append("\t".join([str(cell) if cell is not None else "" for cell in row]))
data_text = "\n".join(data)

# === STEP 2: Draft Outlook mail ===
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

today = datetime.datetime.now().strftime("%Y-%m-%d")
subject = f"{file_name}_{today}"

mail.Subject = subject
mail.Body = f"Dear Team,\n\nHere is the data from {file_name}:\n\n{data_text}\n\nRegards,\nAutomation Bot"

# You can set recipients if needed
# mail.To = "someone@example.com"

# Display (doesn't send yet, just opens draft)
mail.Display()
