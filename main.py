import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    # Setup page
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    # Get filenames
    filename = Path(filepath).stem
    # You could also get invoice number and date from the excel file
    # invoice_nr, date = filename.split("-")

    # Get invoice number and date from filename
    invoice_nr = filename.split("-")[0]        
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1, align="L", border=0)
    
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=14, h=8, txt=f"Date", ln=0, align="L", border=0)
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=50, h=8, txt=f"{date}", ln=1, align="L", border=0)

    pdf.ln(4)

    # Get data from excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Control border size
    border_size = 1

    # Add column headers
    columns = list(df.columns)
    columns = [column.replace("_", " ").title() for column in columns]        

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=border_size)
    pdf.cell(w=50, h=8, txt=columns[1], border=border_size)
    pdf.cell(w=40, h=8, txt=columns[2], border=border_size)
    pdf.cell(w=30, h=8, txt=columns[3], border=border_size)
    pdf.cell(w=30, h=8, txt=columns[4], border=border_size, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"{row['product_id']}", border=border_size)
        pdf.cell(w=50, h=8, txt=f"{row['product_name']}", border=border_size)
        pdf.cell(w=40, h=8, txt=f"{row['amount_purchased']}", border=border_size)
        pdf.cell(w=30, h=8, txt=f"{row['price_per_unit']}", border=border_size)
        pdf.cell(w=30, h=8, txt=f"{row['total_price']}", ln=1, border=border_size)
           
    pdf.output(f"PDFs/{filename}.pdf")
