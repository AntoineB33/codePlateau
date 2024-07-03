import xlwings as xw

# Create a new workbook
wb = xw.Book()
ws = wb.sheets[0]

# Add some data to the sheet
ws.range('A1').value = "Hello"
ws.range('B1').value = "World"

# Save the workbook as .xlsm
wb.save('uzeir\\sample.xlsm')

# Close the workbook
wb.close()

print("Workbook 'sample.xlsm' created successfully.")
