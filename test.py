import xlrd
import os

# Give the location of the file 
loc = ("C:\\Users\\Jesus Covarrubias\\Desktop\\test_xl\\book1.xlsx")

# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)

# For row 0 and column 0 
sheet.cell_value(0, 0)

numOfRows = sheet.nrows
numOfCols = sheet.ncols

# Extracting number of rows and columns
print("Number of rows: " + str(numOfRows))
print("Number of columns: " + str(numOfCols))

if (numOfRows < 0 and numOfCols < 0) or (numOfRows == None and numOfCols ==None):
    
    print("Empty Excel File or Wrong File")
    
else:
    file = open("copy.txt", "w")
    
    for row in range(numOfRows):
        
        #get each row and apply js to it
        file.write("document.getElementById('name').value = " + str(sheet.cell_value(row, 0)) + " \n") 
        file.write("document.getElementById('button').click() \n")
        file.write("document.getElementById('id').value = " + str(sheet.cell_value(row, 0)) + " \n") 
        file.write("document.getElementById('foo').value = " + str(sheet.cell_value(row, 1)) + " \n") 
        file.write("document.getElementById('bar').value = " + str(sheet.cell_value(row, 2)) + " \n") 
        file.write("document.getElementById('foo2').value = " + str(sheet.cell_value(row, 3)) + " \n") 

    file.close() 
