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
    pdf.cell(50, 8, txt=f"Date: {doc_date}")

    pdf.output(f"PDFs/{filename}.pdf")
