import PyPDF2
import io
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import math

pdfmetrics.registerFont(TTFont('PPNM', 'Avancement2020-Thin.ttf'))
def cert_func(name, portfolio, committee):
    existing_pdf = PyPDF2.PdfReader(open("certificate_participation.pdf", "rb"))
    packet = io.BytesIO()
    can1 = canvas.Canvas(packet, pagesize=(842,842))
    can1.setFont(psfontname = "PPNM", size=18)
    if(len(name)<20):
        can1.drawString(365, 275, name)
    else:
        can1.setFont(psfontname = "PPNM", size=14)
        can1.drawString(280, 275, name)
    if(len(portfolio)<20):
        can1.drawString(200, 170, portfolio)
    elif(len(portfolio)>40):
        can1.setFont(psfontname = "PPNM", size=12)
        can1.drawString(125, 170, portfolio)
    else:
        can1.setFont(psfontname = "PPNM", size=14)
        can1.drawString(140, 170, portfolio)
    can1.setFont(psfontname = "PPNM", size=12)
    if(committee == "Lok Sabha"):
        can1.setFont(psfontname = "PPNM", size=18)
        can1.drawString(550, 170, committee)
    elif(committee == "Disarmament and International Security Committee"):
        can1.setFont(psfontname = "PPNM", size=13)
        can1.drawString(450, 170, committee)
    elif(committee == "United Nations Human Rights Council"):
        can1.setFont(psfontname = "PPNM", size=14)
        can1.drawString(450, 170, committee)
    elif(committee == "United Nations Office on Drugs and Crime"):
        can1.setFont(psfontname = "PPNM", size=14)
        can1.drawString(450, 170, committee)
    elif(committee == "Continuous Crisis Committee"):
        can1.setFont(psfontname = "PPNM", size=18)
        can1.drawString(450, 170, committee)
    elif(committee == "United States Senate"):
        can1.setFont(psfontname = "PPNM", size=18)
        can1.drawString(450, 170, committee)
    elif(committee == "United Nations Security Council"):
        can1.setFont(psfontname = "PPNM", size=18)
        can1.drawString(450, 170, committee)
    elif(committee == "International Press Corps"):
        can1.setFont(psfontname = "PPNM", size=18)
        can1.drawString(450, 170, committee)
    can1.save()
    packet.seek(0)

    new_pdf = PyPDF2.PdfReader(packet)
    output = PyPDF2.PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])

    output.add_page(page)

    output_stream = open("Corrections/{name}_{committee}.pdf".format(name = name, committee=committee), "wb")
    output.write(output_stream)
    output_stream.close()

    
    # if(committee == "International Press Corps"):
    #     output_stream = open("IPC/{name}_{committee}.pdf".format(name = name, committee="International Press Corps"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # if(committee == "Disarmament and International Security Committee"):
    #     output_stream = open("DISEC/{name}_{committee}.pdf".format(name = name, committee="Disarmament and International Security Committee"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # elif(committee == "United States Senate"):
    #     output_stream = open("US Senate/{name}_{committee}.pdf".format(name = name, committee="United States Senate"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # elif(committee == "Continuous Crisis Committee"):
    #     output_stream = open("CCC/{name}_{committee}.pdf".format(name = name, committee="Continuous Crisis Committee"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # elif(committee == "United Nations Human Rights Council"):
    #     output_stream = open("UNHRC/{name}_{committee}.pdf".format(name = name, committee="United Nations Human Rights Council"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # elif(committee == "United Nations Office on Drugs and Crime"):
    #     output_stream = open("UNODC/{name}_{committee}.pdf".format(name = name, committee="United Nations Office on Drugs and Crime"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # elif(committee == "United Nations Security Council"):
    #     output_stream = open("UNSC/{name}_{committee}.pdf".format(name = name, committee="United Nations Security Council"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()
    # elif(committee == "Lok Sabha"):
    #     output_stream = open("Lok Sabha/{name}_{committee}.pdf".format(name = name, committee="Lok Sabha"), "wb")
    #     output.write(output_stream)
    #     output_stream.close()

xls = pd.ExcelFile('all_regs.xlsx')

df_corr = pd.read_excel(xls, 'Corrections') 
df1 = pd.read_excel(xls, 'IPC')
df2 = pd.read_excel(xls, 'DISEC')
df3 = pd.read_excel(xls, 'US Senate')
df4 = pd.read_excel(xls, 'CCC')
df5 = pd.read_excel(xls, 'UNHRC')
df6 = pd.read_excel(xls, 'UNODC')
df7 = pd.read_excel(xls, 'UNSC')
df8 = pd.read_excel(xls, 'LS')


all_participants_corr = []
all_participants1 = []
all_participants2 = []
all_participants3 = []
all_participants4 = []
all_participants5 = []
all_participants6 = []
all_participants7 = []
all_participants8 = []


# for index, row in df1.iterrows(): 
#     all_participants1.append({"name":row.Name, "country":row.Portfolio})

# for index, row in df2.iterrows(): 
#     all_participants2.append({"name":row.Name, "country":row.Country})

# for index, row in df3.iterrows(): 
#     all_participants3.append({"name":row.Name, "country":row.Portfolio})

# for index, row in df4.iterrows(): 
#     all_participants4.append({"name":row.Name, "country":row.Portfolio})

# for index, row in df5.iterrows(): 
#     all_participants5.append({"name":row.Name, "country":row.Country})

# for index, row in df6.iterrows(): 
#     all_participants6.append({"name":row.Name, "country":row.Country})

# for index, row in df7.iterrows(): 
#     all_participants7.append({"name":row.Name, "country":row.Country})

# for index, row in df8.iterrows(): 
#     all_participants8.append({"name":row.Name, "country":row.Portfolio})

for index, row in df_corr.iterrows(): 
    all_participants_corr.append({"name":row.Name, "country":row.Portfolio, "committee":row.Committee})


def isNaN(string):
    return string != string

for x in all_participants_corr:
     if(isNaN(x["name"])):
        print("nanhahahah")
     else:
        cert_func(x["name"], x["country"], x["committee"])

# for x in all_participants1:
#     if(isNaN(x["name"])):
#         print("nanhahahah")
#     else:
#         cert_func(x["name"], x["country"], "International Press Corps")
# for x in all_participants2:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "Disarmament and International Security Committee")
# for x in all_participants3:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "United States Senate")
# for x in all_participants4:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "Continuous Crisis Committee")
# for x in all_participants5:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "United Nations Human Rights Council")
# for x in all_participants6:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "United Nations Office on Drugs and Crime")
# for x in all_participants7:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "United Nations Security Council")
# for x in all_participants8:
#      if(isNaN(x["name"])):
#         print("nanhahahah")
#      else:
#         cert_func(x["name"], x["country"], "Lok Sabha")

