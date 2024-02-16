import openpyxl
import csv
from openpyxl.styles import PatternFill

csv_file_path = 'Data'
file_name1 = "Data.csv"
file_name2 = "Data2.csv"
output_file_name = "Data_with_Differences.xlsx"

# Function to read CSV file into a list of lists
def read_csv(file_name):
    data = []
    with open(file_name, 'r', newline='') as csvfile:
        csv_reader = csv.reader(csvfile)
        for row in csv_reader:
            data.append(row)
    return data

# Read CSV files into lists of lists
data1 = read_csv(file_name1)
data2 = read_csv(file_name2)

# Create a new workbook
workbook = openpyxl.Workbook()
sheet = workbook.active

# Write original data from the first CSV file to the Excel sheet
for row_idx, row in enumerate(data1, start=1):
    sheet.append(row)

# Function to highlight differences between two rows
def highlight_diff(row1, row2):
    highlighted_row = []
    for val1, val2 in zip(row1, row2):
        if val1 != val2:
            highlighted_row.append(f"{val1}**")
        else:
            highlighted_row.append(val1)
    return highlighted_row

# Apply highlighting to differences and new data, and write to Excel
for row_idx, (row1, row2) in enumerate(zip(data1, data2), start=1):
    highlighted_row = highlight_diff(row1, row2)
    for col_idx, (value, orig_value) in enumerate(zip(highlighted_row, row1), start=1):
        cell = sheet.cell(row=row_idx, column=col_idx)
        if value != orig_value:
            # Apply yellow fill color to highlight differences
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        if row_idx > len(data1) and value == orig_value:
            # Apply green fill color to highlight new data
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        cell.value = value

# Save the workbook to an Excel file
workbook.save(filename=output_file_name)

print(f"Excel file with differences and new data highlighted written to '{output_file_name}'")
