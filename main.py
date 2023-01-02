import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    # Creates a data frame for each Excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Creates a pdf for each Excel file
    pdf = FPDF(orientation="P", unit='mm', format='A4')

    # Finds the filename for invoice number needed in the pdf
    filepath = Path(filepath).stem
    invoice = filepath.split('-')[0]

    # Retrieves date from filepath
    date = filepath.split('-')[1]

    # Adds page, and adds the invoice number and date on top of the pdf
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0, h=18, txt=f"Invoice Number: {invoice}", border=0, ln=1, align="L")
    pdf.cell(w=0, h=18, txt=f"Date: {date}", border=0, align="L", ln=1)

    # Creates headers for data table
    columns = list(df.columns)
    columns = [item.replace('-', ' ').title() for item in columns]
    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=12, txt=columns[0], border=1, align="C")
    pdf.cell(w=60, h=12, txt=columns[1], border=1, align="C")
    pdf.cell(w=40, h=12, txt=columns[2], border=1, align="C")
    pdf.cell(w=30, h=12, txt=columns[3], border=1, align="C")
    pdf.cell(w=30, h=12, txt=columns[4], border=1, align="C", ln=1)

    # Adds the data table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="B", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=12, txt=str(row["product_id"]), border=1, align="C")
        pdf.cell(w=60, h=12, txt=str(row["product_name"]), border=1, align="C")
        pdf.cell(w=40, h=12, txt=str(row["amount_purchased"]), border=1, align="C")
        pdf.cell(w=30, h=12, txt=str(row["price_per_unit"]), border=1, align="C")
        pdf.cell(w=30, h=12, txt=str(row["total_price"]), border=1, align="C", ln=1)

    # Adds total price in the data and final line
    total_price = df["total_price"].sum()
    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=12, txt="", border=1, align="C")
    pdf.cell(w=60, h=12, txt="", border=1, align="C")
    pdf.cell(w=40, h=12, txt="", border=1, align="C")
    pdf.cell(w=30, h=12, txt="", border=1, align="C")
    pdf.cell(w=30, h=12, txt=str(total_price), border=1, align="C", ln=1)

    final_statement = f"The total price is {total_price}"
    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=16, txt=final_statement, border=0)

    # Outputs a pdf file
    pdf.output(f"PDFs/{filepath}.pdf")





