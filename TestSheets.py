import xlrd
import xlwt
import os.path

#assert os.path.isfile('C/Users/Connor/Documents/TextBook.xlsx')


book = xlrd.open_workbook(r"C:\Users\Connor\Documents\TextBook.xlsx")
sheet = book.sheet_by_index(0)

print(sheet.cell_value(1,1))

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('test')

sheet.write(0,4,10)
workbook.save(r"C:\Users\Connor\Documents\TextBookWrite.xls")