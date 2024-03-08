import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    filename = Path(filepath).stem
    # You could also get invoice number and date from the excel file
    # invoice_nr, date = filename.split("-")

    invoice_nr = filename.split("-")[0]        
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=True, align="L", border=False)
    
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=14, h=8, txt=f"Date", ln=False, align="L", border=0)
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=50, h=8, txt=f"{date}", ln=True, align="L", border=False)
    
    pdf.output(f"PDFs/{filename}.pdf")
