import pandas
import glob #is used to make a list of file paths
from fpdf import FPDF
import openpyxl #it has to be imported to read xlsx files


filepaths = glob.glob(r"excelFiles\*.xlsx")
print(filepaths)

for filepath in filepaths:
    data = pandas.read_excel(filepath,sheet_name="Sheet 1")
    print(data)