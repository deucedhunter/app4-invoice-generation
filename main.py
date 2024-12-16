import pandas as pd
from pathlib import Path
from fpdf import FPDF
import glob

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = str(filename.split("-")[0])
    date = filename.split("-")[1]


    pdf.set_font('Arial', 'B', 16)
    pdf.cell(50, 8, txt=f'Invoice nr. {invoice_nr}', ln=1)

    pdf.set_font('Arial', 'B', 16)
    pdf.cell(50, 8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    column_list = list(df.columns)
    column_list = [column.replace("_", " ").title() for column in column_list]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=8, txt=str(column_list[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(column_list[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(column_list[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(column_list[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(column_list[4]), border=1, ln=1)
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)
    total_sum = str(df['total_price'].sum())
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=total_sum, border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=30, h=8, txt=f"Total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family='Times', size=14, style='B')
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)



    pdf.output("PDFs/" + filename + ".pdf")