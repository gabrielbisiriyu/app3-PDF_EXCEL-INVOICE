import pandas as pd 
from fpdf import FPDF 
from glob import glob 
from pathlib import Path 
filepaths=glob("invoices/*.xlsx")
 
for filepath in filepaths:
    
    pdf=FPDF(orientation='P',unit="mm",format="a4")
    pdf.add_page()

    # To extract only the file number
    filename=Path(filepath).stem 
    invoice_no=filename.split('-')
    date=invoice_no[1]

    pdf.set_font(family="Times",size=17,style="B") 
    pdf.cell(w=50,h=8,txt=f"Invoice no: {invoice_no[0]}",ln=1)  
    pdf.cell(w=50,h=8,txt=f"Date: {date}",ln=1)  
    
    df=pd.read_excel(filepath,sheet_name="Sheet 1")   
    # To get the headers of the xcel files
    headers=list(df.columns) 
    header=[head.replace('_'," ").title() for head in headers]
    pdf.set_font(family="Times",size=12,style="B") 
    pdf.set_text_color(80,80,80) 
    pdf.cell(w=30,h=8,txt=header[0],border=1)
    pdf.cell(w=60,h=8,txt=header[1],border=1) 
    pdf.cell(w=40,h=8,txt=header[2],border=1) 
    pdf.cell(w=30,h=8,txt=header[3] ,border=1)
    pdf.cell(w=25,h=8,txt=header[4],border=1,ln=1) 
        
    for index,row in df.iterrows():
        pdf.set_font(family="Times",size=10) 
        pdf.set_text_color(80,80,80) 
        pdf.cell(w=30,h=8,txt=str(row["product_id"]),border=1)
        pdf.cell(w=60,h=8,txt=row["product_name"],border=1) 
        pdf.cell(w=40,h=8,txt=str(row["amount_purchased"]),border=1) 
        pdf.cell(w=30,h=8,txt=str(row["price_per_unit"]) ,border=1)
        pdf.cell(w=25,h=8,txt=str(row["total_price"]),border=1,ln=1)
        
    total_price=df["total_price"].sum()
    pdf.cell(w=30,h=8,txt="")
    pdf.cell(w=60,h=8,txt="") 
    pdf.cell(w=40,h=8,txt="") 
    pdf.cell(w=30,h=8,txt="")  
    pdf.cell(w=25,h=8,txt=str(total_price),border=1,ln=1)  
    pdf.set_font(family="Times",size=12,style="B")
    pdf.cell(w=70,h=18,txt=f"The total due amount is {total_price} Euros ",ln=1) 
    
    pdf.cell(w=23,h=5,txt="PythonHow")
    pdf.image("pythonhow.png",w=5,h=5) 
    pdf.output(f"PDFs/{filename}.pdf") 

    


