import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
#commit: create minimal pdf for each excel file Sec25

filepaths=glob.glob('invoices/*.xlsx')
print(filepaths)

for filepath in filepaths:
    df=pd.read_excel(filepath,sheet_name='Sheet 1') # excel files need sheet name!!!
    print(df)
    pdf=FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    
    # extract inpice nr from filename
    filename=Path(filepath).stem # get filename !!!
    invoice_nr=filename.split('-')[0]
    # print(f"Invoice nr from filename: {filename} is: {invoice_nr}")

    pdf.set_font(family='Times',size=16,style='B')
    pdf.cell(w=50,h=8,txt=f'Invoice nr.{invoice_nr}')
    pdf.output(f'pdfs/{filename}.pdf')