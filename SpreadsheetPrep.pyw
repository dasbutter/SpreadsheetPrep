import xlrd
import csv
import tkFileDialog
import os
from Tkinter import Tk 
from Tkinter import Label
from decimal import Decimal

#We want to convert a xls to csv file, change the delimeter to a "|" 
#and then strip off any extra spaces in preparation of loading.
#We're also only grabbing the zeroth worksheet name. This can be easily fixed in the future.
def csv_from_excel(path):
    stripped = []
    wb = xlrd.open_workbook(path)
    SheetName = wb.sheet_names()[0]
    sh = wb.sheet_by_name(SheetName)
    filename = path[:-3] + "csv"
    csv_file = open(filename, 'wb')
    wr = csv.writer(csv_file, skipinitialspace=True, delimiter='|',quoting=csv.QUOTE_NONE)

    #Strip off trailing spaces and append it to stripped. 
    #Write each row with new stripped variable.
    #Also handle elements where we do not want to truncate off trailing zeros.    
    PLACES = Decimal(10) ** -4
    for rownum in xrange(sh.nrows):
        for element in range(len(sh.row_values(rownum))):
            if sh.row_values(rownum)[element] in["MIC Antibiotic Dictionary Training Guide","MIC Source Dictionary Training Guide","MIC Organism Dictionary Training Guide"]:
                PLACES = Decimal(10) ** -3
            if isinstance(sh.row_values(rownum)[element],float):
                stripped.append(str(((Decimal(sh.row_values(rownum)[element])).quantize(PLACES))).strip())
            else:
                stripped.append(str(sh.row_values(rownum)[element]).strip(" "))
        wr.writerow(stripped)
        stripped = []            
    csv_file.close()
    
window = Tk()
window.title('LAB Dictionary Spreadsheet Converter')
window.iconbitmap(os.path.normpath(os.getcwd()) + '\icon.ico')
label = Label()
file_path = tkFileDialog.askopenfilename(title='Select Spreadsheet(s) to Load into LAB',multiple=True)
if file_path:
    for f in window.tk.splitlist(file_path):
        f = os.path.normpath(f)
        #f = os.path.normcase(f)
        csv_from_excel(f)
    label = Label(window,text="CSV Conversion is complete." + "\n The Converted file(s) are located in \n" + os.path.dirname(f) + "\nPlease review converted file.")
else:
    label = Label(window,text="No files selected.")
label.pack()

window.mainloop()