import xlwings as xw

print(1)

for book in xw.books:
    if book.fullname == r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\sample2.xlsm":
        wb = book
        is_open = True
        break

# Attach to the workbook by name. Change 'your_workbook_name.xlsm' to your actual workbook name.
# wb = xw.books['sample2.xlsm']

# Save the workbook
wb.save("uzeir\\sample2.xlsm")

# Optional: Specify a path to save the workbook with a different name
# wb.save('path/to/your/filename.xlsm')

# Optional: Save the workbook and close Excel
# wb.save()
# wb.close()
