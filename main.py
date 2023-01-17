import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_page()
    filename = Path(filepath).stem
    inv_nr, date = filename.split("-")
    pdf.set_font("Times", style='B', size=20)
    pdf.cell(w=0, h=12, txt=f"Invoice nr: {inv_nr}", align='L', ln=1, border=1)
    pdf.cell(w=0, h=12, txt=f"Date: {date}", align='L', ln=1, border=1)

    # Reading Excel files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    orig_columns = list(df.columns)
    columns = [c.replace("_", " ").title() for c in orig_columns]

    # Headers
    pdf.set_font(family="Times", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=25, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        # Rest of the cells
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=25, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Adding the total price row
    # x = 10 + (w numbers above except last one) =  170

    pdf.set_font(family="Times", size=10, style='B')
    pdf.set_x(170)
    total_amt = sum(df["total_price"])
    pdf.cell(w=30, h=8, txt=str(total_amt), border=1, ln=1,
             align='L')

    # Adding total sum sentence
    pdf.set_font("Times", style='B', size=12)
    pdf.cell(w=0, h=10, txt=f"Total amount is ${total_amt}.", align='L', ln=1)

    # Adding company logo
    pdf.set_font("Times", style='B', size=14)
    pdf.cell(w=25, h=10, txt=f"PythonHow", align='L')
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}_PDF.pdf")
