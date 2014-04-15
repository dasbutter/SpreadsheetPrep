import xlrd
import csv
import tkFileDialog
from Tkinter import Tk 
from Tkinter import Label

#We want to convert a xls to csv file, change the delimeter to a "|" 
#and then strip off any extra spaces in preparation of loading.
def csv_from_excel(path):
    stripped = []
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_name('Sheet1')
    filename = path[:-3] + "csv"
    csv_file = open(filename, 'wb')
    wr = csv.writer(csv_file, skipinitialspace=True, delimiter='|',quoting=csv.QUOTE_NONE)

    #Strip off trailing spaces and append it to stripped. 
    #Write each row with new stripped variable    
    for rownum in xrange(sh.nrows):
        for element in range(len(sh.row_values(rownum))):
            stripped.append(sh.row_values(rownum)[element].strip())
        wr.writerow(stripped)
        stripped = []
            
    csv_file.close()

window= Tk()
file_path = tkFileDialog.askopenfilename(multiple=True)    
if file_path:
    for f in file_path.split(' '):
        csv_from_excel(f)
        label = Label(window,text="CSV Conversion is complete for " + file_path)
label.pack()

window.mainloop()