import openpyxl
import shelve
file = shelve.open('data')
wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()
sheet.title = '2017'
tabcount = 0
for table in file['2017']:
    count = 0
    for row in table:
        count = count + 1
        for i in range(len(row)):
            sheet.cell(row=count+tabcount, column=i+1, value=row[i])
    tabcount = tabcount + len(table)
wb.save('1.xlsx')
file.close()
