import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_page()
    filename = Path(filepath).stem
    inv_nr, date = filename.split("-")
    pdf.set_font("Times", style='B', size=20)
    pdf.cell(w=0, h=12, txt=f"Invoice nr: {inv_nr}", align='L', ln=1, border=1)
    pdf.cell(w=0, h=12, txt=f"Date: {date}", align='L', ln=1, border=1)

    pdf.cell(w=0, h=12, align='L', ln=1, border=1)

    pdf.output(f"PDFs/{filename}_PDF.pdf")
