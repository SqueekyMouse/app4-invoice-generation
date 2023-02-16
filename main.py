import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
#commit: add inv date hdr table rows logo Sec26

def add_table(pdf,table=['','','','',''],fsize=10,fstyle=''):
    pdf.set_font(family='Times',size=fsize,style=fstyle)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=25,h=8,txt=table[0],border=1)
    pdf.cell(w=70,h=8,txt=table[1],border=1)
    pdf.cell(w=35,h=8,txt=table[2],border=1)
    pdf.cell(w=30,h=8,txt=table[3],border=1)
    pdf.cell(w=30,h=8,txt=table[4],border=1)
    pdf.ln()
    return(pdf)

filepaths=glob.glob('invoices/*.xlsx')
# print(filepaths)

for filepath in filepaths:
    pdf=FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    
    # extract invoice nr and date from filename
    filename=Path(filepath).stem # get filename !!!
    invoice_nr,inv_date=filename.split('-') # list unpacking!!!

    pdf.set_font(family='Times',size=16,style='B')
    pdf.cell(w=50,h=8,txt=f'Invoice nr.{invoice_nr}',ln=1)

    pdf.set_font(family='Times',size=16,style='B')
    pdf.cell(w=50,h=8,txt=f'Date {inv_date}',ln=1)

    df=pd.read_excel(filepath,sheet_name='Sheet 1') # excel files need sheet name!!!
    # print(df)

    # add header
    # columns=list(df.columns) # get headers or col list!!!
    # df.columns is an Index which is an iterable so no need to conv to list!!!
    columns=[i.replace('_',' ').title() for i in df.columns] #list comprehension!!!
    pdf=add_table(pdf,table=columns,fstyle='B')

    # add rows
    for index,row in df.iterrows():
        frow=[str(i) for i in row]
        pdf=add_table(pdf,table=frow)
    
    total=df['total_price'].sum()
    pdf=add_table(pdf,table=['','','','',str(total)])

    # add total sum sentence
    pdf.ln(20)
    pdf.set_font(family='Times',size=12,style='B')
    pdf.set_text_color(0,0,0)
    pdf.cell(w=30,h=10,txt=f"The total amount due is {df['total_price'].sum()} Euros",ln=1)

    # add company name and logo
    pdf.set_font(family='Times',size=14,style='B')
    pdf.cell(w=35,h=10,txt='Python Howto')
    pdf.image('pythonhow.png',w=10)

    pdf.output(f'pdfs/{filename}.pdf')