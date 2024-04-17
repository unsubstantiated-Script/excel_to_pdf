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
    invoice_nr, doc_date = filename.split("-")

    pdf.set_font('Times', 'B', size=16)
    pdf.cell(50, 8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font('Times', 'B', size=16)
    pdf.cell(50, 8, txt=f"Date: {doc_date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    columns = df.columns
    # Looping, removing underscores, and capitalizing
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font('Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 8, txt=columns[0], border=1)
    pdf.cell(70, 8, txt=columns[1], border=1)
    pdf.cell(30, 8, txt=columns[2], border=1)
    pdf.cell(30, 8, txt=columns[3], border=1)
    pdf.cell(30, 8, txt=columns[4], border=1, ln=1)

    # Print values in body of cells.
    for index, row in df.iterrows():
        pdf.set_font('Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(30, 8, txt=f"{row['product_id']}", border=1)
        pdf.cell(70, 8, txt=f"{row['product_name']}", border=1)
        pdf.cell(30, 8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(30, 8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(30, 8, txt=f"{row['total_price']}", border=1, ln=1)

    # Space out and add total sum
    total_sum = df['total_price'].sum()
    pdf.set_font('Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 8, txt="", border=1)
    pdf.cell(70, 8, txt="", border=1)
    pdf.cell(30, 8, txt="", border=1)
    pdf.cell(30, 8, txt="", border=1)
    pdf.cell(30, 8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font('Times', size=10, style='B')
    pdf.cell(30, 8, txt=f"The total price is {total_sum}", ln=1)

    # Add Company Name and Logo
    pdf.set_font('Times', size=14, style='B')
    pdf.cell(25, 8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
