import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    # Creates a data frame for each Excel file
    df = pd.read_excel(filepath)
    # Creates a pdf for each Excel file
    pdf = FPDF(orientation="P", unit='mm', format='A4')

    # Finds the filename for invoice number needed in the pdf
    filepath = Path(filepath).stem
    invoice = filepath.split('-')[0]

    # Adds page, and adds the invoice number on top of the pdf
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0, h=18, txt=f"Invoice Number: {invoice}", border=0, ln=1, align="l")

    # Outputs a pdf file
    pdf.output(f"PDFs/{filepath}.pdf")





