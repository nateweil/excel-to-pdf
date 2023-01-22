import pandas as pd
from fpdf import FPDF
import glob
import openpyxl
from pathlib import Path

#Will return a list object with all .xlsx files in the folder.
filepaths = glob.glob("invoices/*.xlsx")


for path in filepaths:
    data = pd.read_excel(path, sheet_name="Sheet 1")

    name = Path(path).stem.split("-")[0]

    pdf = FPDF(orientation="Portrait", unit="mm", format="letter")
    pdf.add_page()

    pdf.set_font(family="Arial", size=20, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice #{name}")
    pdf.output(f"PDFs/{name}.pdf")