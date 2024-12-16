import pandas as pd
from pathlib import Path
from fpdf import FPDF
import glob

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(50, 8, txt=f'Invoice nr. {filename.split("-")[0]}')

    pdf.output("PDFs/" + filename + ".pdf")