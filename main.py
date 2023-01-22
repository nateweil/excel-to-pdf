import pandas as pd
import fpdf as pdf
import glob
import openpyxl

#Will return a list object with all .xlsx files in the folder.
filepaths = glob.glob("invoices/*.xlsx")


for path in filepaths:
    data = pd.read_excel(path, sheet_name="Sheet 1")
    print(data)