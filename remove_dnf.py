import openpyxl

def delete_dnf_rows(file_path):
    # Load the workbook and select the active sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    rows_deleted = 0
    row = sheet.max_row

    # Loop backwards through the rows to avoid skipping rows after deletion
    while row >= 1:
        if sheet[f'A{row}'].value == "DNF":
            sheet.delete_rows(row)
            rows_deleted += 1
        row -= 1

    # Save the modified workbook
    workbook.save(file_path)
    
    # Print the number of rows deleted
    print(f"Total rows deleted: {rows_deleted}")

# Path to your Excel file
file_path = 'modified_eyemark_budget.xlsx'

# Call the function
delete_dnf_rows(file_path)

