'''
Final year Project interface
'''
#All the libraries required for various purposes
from tkinter import *
from PIL import ImageTk, Image
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.text
from matplotlib import patches
from matplotlib.backends.backend_pdf import PdfPages
import fitz
from fitz.utils import insertImage

#importing from excel
Panel_data = pd.read_excel("D:\Ken related\GIT\Git again\Haz_cum.xlsm","Panel_Data")
Panel_data_headings = list(Panel_data)
Precaution = pd.read_excel("D:\Ken related\GIT\Git again\\Haz_cum.xlsm","Precaution",index_col=0, header = 0)
Precaution_headings = list(Precaution)
Mitigation = pd.read_excel("D:\Ken related\GIT\Git again\\Haz_cum.xlsm","Mitigation", index_col=0, header = 0)
Mitigation_headings = list(Mitigation)

#Headers for the pdf files
def headerInfoLow(n,d,l,z):
        plt.annotate("Name of the worker: "+n, xy=(-90, 500), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Designation of the worker: "+d, xy=(-90, 480), xycoords='axes points',fontsize=20, weight="bold")        
        plt.annotate("Location of the worker: "+ "Panel "+ str(l), xy=(-90, 460), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Zone of comfortability: "+ z, xy=(-90,440), xycoords='axes points',fontsize=20, weight="bold")       
        plt.annotate("Golden Rules", xy=(-90, 400), xycoords='axes points',fontsize=20, weight = "bold") 
        if d!="Common Worker":
            plt.annotate("Definition of Responsibility", xy=(-90,140), xycoords='axes points',fontsize=20, weight="bold")

def headerInfoMed(n,d,l,z):
        plt.annotate("Name of the worker: "+n, xy=(-90, 500), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Designation of the worker: "+d, xy=(-90, 480), xycoords='axes points',fontsize=20, weight="bold")        
        plt.annotate("Location of the worker: "+ "Panel "+ str(l), xy=(-90, 460), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Zone of comfortability: "+ z, xy=(-90,440), xycoords='axes points',fontsize=20, weight="bold")       
        plt.annotate("Golden Rules", xy=(-90, 400), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Mitigation Methods", xy=(-90,180), xycoords='axes points',fontsize=20, weight="bold")
        plt.annotate("Frugal Techniques for roof fall detection", xy=(-90,50), xycoords='axes points',fontsize=20, weight="bold")

def headerInfoHigh(n,d,l,z):
        plt.annotate("Name of the worker: "+n, xy=(-90, 500), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Designation of the worker: "+d, xy=(-90, 480), xycoords='axes points',fontsize=20, weight="bold")        
        plt.annotate("Location of the worker: "+ "Panel "+ str(l), xy=(-90, 460), xycoords='axes points',fontsize=20, weight="bold") 
        plt.annotate("Zone of comfortability: "+ z, xy=(-90,440), xycoords='axes points',fontsize=20, weight="bold")      
        plt.annotate("Mitigation Methods", xy=(-90,400), xycoords='axes points',fontsize=20, weight="bold")
        if d!="Common Worker":
            plt.annotate("Designation specific mitigation", xy=(-90,230), xycoords='axes points',fontsize=20, weight="bold")
            plt.annotate("Frugal Techniques for roof fall detection", xy=(-90,100), xycoords='axes points',fontsize=20, weight="bold")
        if d=="Common Worker":
            plt.annotate("Frugal Techniques for roof fall detection", xy=(-90,230), xycoords='axes points',fontsize=20, weight="bold")


def headerInfoVeryHigh(n,d,l,z):
        plt.annotate("Name of the worker: "+n, xy=(-90, 500), xycoords='axes points',fontsize=18, weight="bold") 
        plt.annotate("Designation of the worker: "+d, xy=(-90, 480), xycoords='axes points',fontsize=18, weight="bold")        
        plt.annotate("Location of the worker: "+"Panel "+ str(l), xy=(-90, 460), xycoords='axes points',fontsize=18, weight="bold") 
        plt.annotate("Zone of comfortability: "+ z, xy=(-90,440), xycoords='axes points',fontsize=18, weight="bold")       
        plt.annotate("Mitigation Methods", xy=(-90, 400), xycoords='axes points',fontsize=18, weight="bold") 
        plt.annotate("If Trapped", xy=(-90,260), xycoords='axes points',fontsize=18, weight="bold")
        if d!="Common Worker":
            plt.annotate("Designation specific Mitigation", xy=(-90,140), xycoords='axes points',fontsize=18, weight="bold")

def zoneOfC(risk):
    if risk<28:
        return "Green"
    elif risk<48:
        return "Yellow"
    elif risk<70:
        return "Orange"
    else:
        return "Red"

def view_report():

    name_ = ename.get()
    desg = desg_.get()
    loca_ = elocation.get()
    loca_ = int(loca_[-1])
    risk_rating = Panel_data[Panel_data_headings[-1]][loca_-1]
    zone_ = zoneOfC(risk_rating)


    if desg=="Overman":
            if risk_rating <28:
                # with PdfPages(name_ + '.pdf') as pdf:
                plt.figure(figsize=(20,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoLow(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                for x in range(1,7):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[3]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                plt.show()

            elif risk_rating <48:
                plt.figure(figsize=(17,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoMed(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                plt.show()   

            elif risk_rating <70:
                # Page 1
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                for x in range(2,5):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[21]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=14)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                plt.show()

            # Very high risk
            else:
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoVeryHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                for x in range(1,5):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                for x in range(2,5):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[21]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=14)
                plt.show()

    if desg=="Mining Sirdir":
            if risk_rating <28:
                plt.figure(figsize=(20,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoLow(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                for x in range(1,7):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[5]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                plt.show()

            elif risk_rating <48:
                plt.figure(figsize=(17,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoMed(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                plt.show()
                
            elif risk_rating <70:
                # Page 1
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                for x in range(2,5):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[23]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=11)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                plt.show()

            # Very high risk
            else:
                plt.figure(figsize=(20,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoVeryHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                for x in range(1,5):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                for x in range(2,5):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[23]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=13)
                plt.show()

    if desg=="Shotfirer":
            if risk_rating <28:
                plt.figure(figsize=(20,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoLow(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                for x in range(1,7):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[7]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                plt.show()

            elif risk_rating <48:
                plt.figure(figsize=(17,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoMed(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                plt.show()
                
            elif risk_rating <70:
                # Page 1
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                for x in range(2,6):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[25]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=14)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                plt.show()

            # Very high risk
            else:
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoVeryHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                for x in range(1,5):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                for x in range(2,6):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[25]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=12)
                plt.show()

    if desg=="Timberman":
            if risk_rating <28:
                plt.figure(figsize=(20,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoLow(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                for x in range(1,7):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[9]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                plt.show()

            elif risk_rating <48:
                plt.figure(figsize=(17,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoMed(name_,desg,loca_,zone_) 
                for x in range(1,10):
                    plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                plt.show()
                
            elif risk_rating <70:
                # Page 1
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                for x in range(2,5):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[27]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=14)
                for x in range(1,4):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                plt.show()

            # Very high risk
            else:
                plt.figure(figsize=(15,10))
                plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                plt.axis('off')
                headerInfoVeryHigh(name_,desg,loca_,zone_) 
                for x in range(1,6):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                for x in range(1,5):
                    plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                for x in range(2,5):
                    plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[27]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=14)
                plt.show()

def generate_report():
    name_ = ename.get()
    ename.delete(0,END)
    desg = desg_.get()
    loca_ = elocation.get()
    loca_ = int(loca_[-1])
    elocation.delete(0,END)
    risk_rating = int(Panel_data[Panel_data_headings[-1]][loca_-1])
    zone_ = zoneOfC(risk_rating)


    if desg=="Overman":
            if risk_rating <28:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(20,10))
                    border = plt.figure(figsize=(20,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoLow(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                    for x in range(1,7):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[3]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                    pdf.savefig(edgecolor=border.get_edgecolor())

            elif risk_rating <48:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(17,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoMed(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                
            elif risk_rating <70:
                with PdfPages(name_ + '.pdf') as pdf:
                    # Page 1
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                    for x in range(2,5):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[21]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=14)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # Page 2
                    plt.figure(figsize=(10,11))   
                    border = plt.figure(figsize=(10,11), linewidth=10, edgecolor="#8B4513")                 
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
                #opening the file again to add an image on Page 2
                doc = fitz.open(name_ + '.pdf')
                rect= fitz.Rect(50,50,700,700)
                page = doc.loadPage(1)
                page.insertImage(rect, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                doc.saveIncr()
            # Very high risk
            else:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoVeryHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                    for x in range(1,5):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                    for x in range(2,5):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[21]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # plt.close()
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.annotate("HAZARD REPORT FORM", xy=(490,570), xycoords='axes points',fontsize=27, weight="bold")
                    
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
             #opening the file again to add an image at the top
                doc = fitz.open(name_ + '.pdf')
                rect1= fitz.Rect(-500,30,500,650)
                rect2 = fitz.Rect(500,100,1100,580)
                page = doc.loadPage(1)
                page.insertImage(rect1, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                page.insertImage(rect2, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\annotation.jpg")
                doc.saveIncr()
                
    if desg=="Mining Sirdar":
            if risk_rating <28:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(20,10))
                    border = plt.figure(figsize=(20,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoLow(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                    for x in range(1,7):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[5]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                    pdf.savefig(edgecolor=border.get_edgecolor())

            elif risk_rating <48:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(17,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoMed(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                
            elif risk_rating <70:
                with PdfPages(name_ + '.pdf') as pdf:
                    # Page 1
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                    for x in range(2,5):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[23]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=11)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # Page 2
                    plt.figure(figsize=(10,11))   
                    border = plt.figure(figsize=(10,11), linewidth=10, edgecolor="#8B4513")                 
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
                #opening the file again to add an image on Page 2
                doc = fitz.open(name_ + '.pdf')
                rect= fitz.Rect(50,50,700,700)
                page = doc.loadPage(1)
                page.insertImage(rect, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                doc.saveIncr()

            # Very high risk
            else:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(20,10))
                    border = plt.figure(figsize=(20,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoVeryHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                    for x in range(1,5):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                    for x in range(2,5):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[23]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=13)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # plt.close()
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.annotate("HAZARD REPORT FORM", xy=(490,570), xycoords='axes points',fontsize=27, weight="bold")
                    
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
             #opening the file again to add an image at the top
                doc = fitz.open(name_ + '.pdf')
                rect1= fitz.Rect(-500,30,500,650)
                rect2 = fitz.Rect(500,100,1100,580)
                page = doc.loadPage(1)
                page.insertImage(rect1, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                page.insertImage(rect2, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\annotation.jpg")
                doc.saveIncr()
    if desg=="Shotfirer":
            if risk_rating <28:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(20,10))
                    border = plt.figure(figsize=(20,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoLow(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                    for x in range(1,7):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[7]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                    pdf.savefig(edgecolor=border.get_edgecolor())

            elif risk_rating <48:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(17,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoMed(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                
            elif risk_rating <70:
                with PdfPages(name_ + '.pdf') as pdf:
                    # Page 1
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                    for x in range(2,6):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[25]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=14)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # Page 2
                    plt.figure(figsize=(10,11))   
                    border = plt.figure(figsize=(10,11), linewidth=10, edgecolor="#8B4513")                 
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
                #opening the file again to add an image on Page 2
                doc = fitz.open(name_ + '.pdf')
                rect= fitz.Rect(50,50,700,700)
                page = doc.loadPage(1)
                page.insertImage(rect, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                doc.saveIncr()

            # Very high risk
            else:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoVeryHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                    for x in range(1,5):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                    for x in range(2,6):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[25]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=12)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # plt.close()
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.annotate("HAZARD REPORT FORM", xy=(490,570), xycoords='axes points',fontsize=27, weight="bold")
                    
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
             #opening the file again to add an image at the top
                doc = fitz.open(name_ + '.pdf')
                rect1= fitz.Rect(-500,30,500,650)
                rect2 = fitz.Rect(500,100,1100,580)
                page = doc.loadPage(1)
                page.insertImage(rect1, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                page.insertImage(rect2, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\annotation.jpg")
                doc.saveIncr()
    if desg=="Timberman":
            if risk_rating <28:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(20,10))
                    border = plt.figure(figsize=(20,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoLow(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                    for x in range(1,7):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[9]][x+1], xy=(-90, 140-x*25), xycoords='axes points',fontsize=12)
                    pdf.savefig(edgecolor=border.get_edgecolor())

            elif risk_rating <48:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(17,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoMed(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                
            elif risk_rating <70:
                with PdfPages(name_ + '.pdf') as pdf:
                    # Page 1
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                    for x in range(2,5):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[27]][x], xy=(-90, 230-(x-1)*25), xycoords='axes points',fontsize=14)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 100-x*25), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # Page 2
                    plt.figure(figsize=(10,11))   
                    border = plt.figure(figsize=(10,11), linewidth=10, edgecolor="#8B4513")                 
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
                #opening the file again to add an image on Page 2
                doc = fitz.open(name_ + '.pdf')
                rect= fitz.Rect(50,50,700,700)
                page = doc.loadPage(1)
                page.insertImage(rect, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                doc.saveIncr()

            # Very high risk
            else:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoVeryHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                    for x in range(1,5):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                    for x in range(2,5):
                        plt.annotate(str(x-1)+") "+Mitigation[Mitigation_headings[27]][x], xy=(-90, 160-x*20), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # plt.close()
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.annotate("HAZARD REPORT FORM", xy=(490,570), xycoords='axes points',fontsize=27, weight="bold")
                    
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
             #opening the file again to add an image at the top
                doc = fitz.open(name_ + '.pdf')
                rect1= fitz.Rect(-500,30,500,650)
                rect2 = fitz.Rect(500,100,1100,580)
                page = doc.loadPage(1)
                page.insertImage(rect1, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                page.insertImage(rect2, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\annotation.jpg")
                doc.saveIncr()

    if desg=="Common Worker":
            if risk_rating <28:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(20,10))
                    border = plt.figure(figsize=(20,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoLow(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*25), xycoords='axes points',fontsize=12)
                    pdf.savefig(edgecolor=border.get_edgecolor())

            elif risk_rating <48:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(17,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoMed(name_,desg,loca_,zone_) 
                    for x in range(1,10):
                        plt.annotate(str(x)+") "+Precaution[Precaution_headings[0]][x], xy=(-90, 400-x*20), xycoords='axes points',fontsize=10)
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90, 180-x*20), xycoords='axes points',fontsize=13)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 50-x*20), xycoords='axes points',fontsize=13)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                
            elif risk_rating <70:
                with PdfPages(name_ + '.pdf') as pdf:
                    # Page 1
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(17,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*25), xycoords='axes points',fontsize=14)
                    for x in range(1,4):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[17]][x], xy=(-90, 230-x*25), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # Page 2
                    plt.figure(figsize=(10,11))   
                    border = plt.figure(figsize=(10,11), linewidth=10, edgecolor="#8B4513")                 
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
                #opening the file again to add an image on Page 2
                doc = fitz.open(name_ + '.pdf')
                rect= fitz.Rect(50,50,700,700)
                page = doc.loadPage(1)
                page.insertImage(rect, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                doc.saveIncr()

            # Very high risk
            else:
                with PdfPages(name_ + '.pdf') as pdf:
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.title("HAZCOM REPORT FOR ROOF FALLS", fontsize= 30, weight= "bold")
                    plt.axis('off')
                    headerInfoVeryHigh(name_,desg,loca_,zone_) 
                    for x in range(1,6):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[0]][x], xy=(-90,400-x*20), xycoords='axes points',fontsize=14)
                    for x in range(1,5):
                        plt.annotate(str(x)+") "+Mitigation[Mitigation_headings[3]][x], xy=(-90, 260-x*20), xycoords='axes points',fontsize=14)
                    pdf.savefig(edgecolor=border.get_edgecolor())
                    # plt.close()
                    plt.figure(figsize=(15,10))
                    border = plt.figure(figsize=(15,10), linewidth=10, edgecolor="#8B4513")
                    plt.annotate("HAZARD REPORT FORM", xy=(490,570), xycoords='axes points',fontsize=27, weight="bold")
                    
                    plt.axis('off')
                    pdf.savefig(edgecolor=border.get_edgecolor())
             #opening the file again to add an image at the top
                doc = fitz.open(name_ + '.pdf')
                rect1= fitz.Rect(-500,30,500,650)
                rect2 = fitz.Rect(500,100,1100,580)
                page = doc.loadPage(1)
                page.insertImage(rect1, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\emergency.jpg")
                page.insertImage(rect2, filename="D:\\Ken related\\HazCom_students\\Project\\interface\\annotation.jpg")
                doc.saveIncr()

root = Tk()
root.geometry("300x500")

#Icon of the program
root.iconbitmap("D:\Ken related\HazCom_students\Project\interface\\biohazard.ico")

#Title of the Program
root.title("HazCom for Roof falls")

#Frame for the program
frame = LabelFrame(root, text= "Enter details as suggested", padx = 20, pady = 20)
frame.pack(padx = 20, pady = 20)

#Labels and input boxes

worker_name = Label(frame, text ="Enter your name: ")
ename = Entry(frame, width = 30, borderwidth = 5)
empty1 = Label(frame, text = " ")
desg_ = StringVar()
designation = Label(frame, text ="Enter you Designation: ")
edesignation = OptionMenu(frame, desg_, "Overman", "Mining Sirdar", "Shotfirer", "Timberman", "Common Worker")
empty2 = Label(frame, text = " ")
location = Label(frame, text ="Enter the Location: ")
elocation = Entry(frame,width = 30, borderwidth = 5 )
empty3 = Label(frame, text = " ")

#Report generating button

view_button = Button(frame, text = "View the Report", command = lambda: view_report(), padx =10, pady = 10)
empty4 = Label(frame, text = " ")
report_button = Button(frame, text = "Generate Report", command= lambda: generate_report(),  padx =10, pady = 10)
empty5 = Label(frame, text = " ")
exit_buttton  = Button(frame, text = "Exit Program", command = root.quit , padx =10, pady = 10)

#Using Grid

# worker_name.grid(row=0, column = 0)
# ename.grid(row=1, column=0,columnspan = 1)
# empty1.grid(row =2 , column =0 )
# designation.grid(row=3, column = 0)
# edesignation.grid(row=4, column=0,columnspan = 1)
# # empty2.pack()
# location.grid(row=5, column = 0)
# elocation.grid(row=6, column=0,columnspan = 1)
# empty3.grid(row= 7, column =0 )
# view_button.grid(row=8 , column =0 )
# empty4.grid(row=9 , column = 0)
# report_button.grid(row= 8, column = 1)
# empty5.grid(row= 11, column =0 )
# exit_buttton.grid(row=12 , column =0 )

#Using pack
worker_name.pack()
ename.pack()
empty1.pack()
designation.pack()
edesignation.pack()
empty2.pack()
location.pack()
elocation.pack()
empty3.pack()
view_button.pack()
empty4.pack()
report_button.pack()
empty5.pack()
exit_buttton.pack()


root.mainloop()