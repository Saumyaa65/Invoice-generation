import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

filepaths=glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    df=pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf=FPDF(orientation="P", unit="mm", format="A4")
    filename=Path(filepath).stem
    filename=filename.split("-")
    in_nr=filename[0]
    date=filename[1]
    pdf.add_page()
    pdf.set_font(family="Arial", style="BU", size=16)
    pdf.cell(w=0, h=20, txt=f"Invoice nr. {in_nr}", align="C", ln=1)
    pdf.output(f"PDFs/{filename}.pdf")