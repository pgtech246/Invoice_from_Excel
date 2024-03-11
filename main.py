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
    pdf.cell(w=14, h=8, txt=f"Date", ln=False, align="L", border=0)
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=50, h=8, txt=f"{date}", ln=True, align="L", border=0)

    pdf.ln(4)

    # Get data from excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add column headers
    columns = list(df.columns)
    columns = [column.replace("_", " ").title() for column in columns]        

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=50, h=8, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=40, h=8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['total_price']}", ln=1, border=1)

    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)

    # Calculate the sum of the total prices
    sum_of_total_prices = df['total_price'].sum()

    pdf.cell(w=30, h=8, txt=f"{sum_of_total_prices}", border=1, ln=1)

    pdf.ln(4)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"The total due amount is {sum_of_total_prices} Euros.", ln=1, align="L", border=0)
    pdf.cell(w=50, h=8, txt=f"PythonHow", align="L", border=0)
    
    # Add image
    image_path = "pythonhow.png"
    x = pdf.get_x() - 27
    y = pdf.get_y() + 2
    pdf.image(image_path, x=x, y=y, w=5)

    pdf.output(f"PDFs/{filename}.pdf")
