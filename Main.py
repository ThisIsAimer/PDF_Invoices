import pandas
import glob #is used to make a list of file paths
from fpdf import FPDF
import openpyxl #it has to be imported to read xlsx files
import pathlib


filepaths = glob.glob(r"excelFiles\*.xlsx")
print(filepaths)

for filepath in filepaths:
    data = pandas.read_excel(filepath,sheet_name="Sheet 1")
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    filename= pathlib.Path(filepath).stem #shows exact filename
    invoice_number = filename.split("-")[0]
    pdf.add_page()
    pdf.set_font(family="times",style="B",size=15)
    pdf.cell(w=50,h=8,txt=f"Invoice number: {invoice_number}",ln=1)

    pdf.output(fr"invoicePDFS\{filename}.pdf")
    print(data)