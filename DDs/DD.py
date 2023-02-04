import pandas as pd
import numpy as np

from datetime import timedelta
from datetime import time
from datetime import date
from datetime import datetime
import time as t
import locale
from datetime import date
import pathlib

import os
import glob
from openpyxl import load_workbook
from styleframe import StyleFrame
from styleframe import Styler

import re
import PyPDF2
import warnings
import shutil

import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
import shutil

from tkinter import * 
import sys
from tkcalendar import DateEntry
from tkinter import ttk

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import StaleElementReferenceException
import chromedriver_autoinstaller

warnings.filterwarnings("ignore")

class DD():

    package_directory = os.getcwd()
            
    def __init__(self, path_driver):
        self.path_driver = path_driver
        #Update chromedriver and put it in path
        chromedriver_autoinstaller.install()
        #Check if the needed folders already exist or not
        language = locale.getdefaultlocale()[0]
        FR_desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Bureau')
        EN_desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        if language == "fr_FR":
            desktop_path = FR_desktop_path
            newpath = FR_desktop_path +"\\FR_dd"
            #DD folder check
            if not os.path.exists(newpath):
                os.makedirs(newpath)
                print(f"New DD folder created at this location: {newpath} ")
                self.path = newpath
            else :
                self.path = newpath
        else:
            newpath = EN_desktop_path +"\\FR_dd"
            if not os.path.exists(newpath):
                os.makedirs(self.newpath)
                print(f"New folder created at this location: {newpath} ")
                self.path = newpath

    def extract_DD(self):
        
        #Function to extract the information from the pdfs and store the cleaned data into an excel file
        np.random.seed(1)
        df = pd.DataFrame(columns=[
            "PRIORITY",
            "REGISTRANT",
            "COMPANY",
            "OPERATION",
            "QUANTITY",
            "AVG PRICE",
            "VALUE",
            "TRADE DATE",
            "DISCLOSING DATE",
            "Nature de la déclaration",
            "Commentaire AMF",
            "Référence AMF",
        ]
        )

        priority_list = ["Bernard Arnault",
                        "Arnaud Lagardere",
                        "Arnaud Lagardère",
                        "Vincent Bolloré",
                        "Vincent Bollore",
                        "Xavier Niel",
                        "Amber Capital",
                        "Christian Dior",
                        "Vivendi",
                        "Bolloré",
                        "Bollore"]


        directory = self.path 
        for file in os.listdir(directory + "\\"):
            filename = os.fsdecode(file)
            print(filename)

            reader = PyPDF2.PdfFileReader(directory + "\\"+ filename)

            texte_utilisable = reader.getPage(0).extractText()
            count = reader.numPages

            for i in range(1, count):
                page = reader.getPage(i)
                texte_utilisable += page.extractText()
            texte_util = (texte_utilisable + '.')[:-1]
            A= texte_util

            prices = []
            volumes = []

            for occurrence in range (texte_utilisable.count("INFORMATIONS AGREGEES")):
                prices.append(float(re.search("PRIX : [0-9](.*?)\.[0-9]{4}", texte_util[texte_util.find("INFORMATIONS AGREGEES"):],).group(0).replace(" ", "")[5:]))
                volumes.append(float(re.search("VOLUME : [0-9](.*?)\.[0-9]{4}", texte_util[texte_util.find("INFORMATIONS AGREGEES"):],).group(0).replace(" ", "")[7:]))
                texte_util = texte_util[texte_util.find("INFORMATIONS AGREGEES") + len("INFORMATIONS AGREGEES"):]
            if float(sum(prices) + sum(volumes)) != 0.:
                price = np.average(prices, weights=volumes)
                volume = sum(volumes)
                value = price * volume
            else:
                value = 0
            for i in prices:
                a=0
                if i == 0:
                    volume = volumes[a]
                    value =prices[a]
                    a+=1

            référence_AMF = A[:12]
            commentaire_AMF=texte_util[texte_util.find("COMMENTAIRES :")+len("COMMENTAIRES :"):texte_util.find("Les données à caractère personnel")]

            exercer = texte_utilisable[texte_utilisable.find("ETROITEMENT LIEE :") + len("ETROITEMENT LIEE :"): texte_utilisable.find("NOTIFICATION INITIALE")]
            company_name = texte_utilisable[texte_utilisable.find("NOM :") + len("NOM : "): texte_utilisable.find("DETAIL DE LA TRANSACTION")]
            print(company_name)

            if company_name.find("LEI :") != -1:
                company_name = company_name[:company_name.find("LEI :")]
            description = texte_utilisable[texte_utilisable.find("DESCRIPTION DE L™INSTRUMENT FINANCIER :") + len("DESCRIPTION DE L™INSTRUMENT FINANCIER : "): texte_utilisable.find("INFORMATION DETAILLEE PAR OPERATION")]
            if texte_utilisable.find("CODE D™IDENTIFICATION DE L™INSTRUMENT FINANCIER") != -1:
                description = description[:description.find("CODE D™IDENTIFICATION DE L™INSTRUMENT FINANCIER")]
            if description == "Action":
                description = texte_utilisable[texte_utilisable.find("NATURE DE LA TRANSACTION : ") + len("NATURE DE LA TRANSACTION : "): texte_utilisable.find("DESCRIPTION DE L™INSTRUMENT FINANCIER :")]
            check_notif_ou_modif = texte_utilisable[texte_utilisable.find("NOTIFICATION INITIALE / MODIFICATION:"): texte_utilisable.find("COORDONNEES DE L™EMETTEUR")]
            if "Notification initiale" in check_notif_ou_modif:
                nature = "Notification initiale"
            else :
                nature = "Modification !"

            description_transfo = {
                "Acquisition": "Buy",
                "Cession": "Sell",
                "Exercise": "Exercice",
                "Souscription": "Subscription",
                "Nantissement d'un compte titres": "Pledge of securities",
                "Nantissement de titres": "Pledge of securities"
            }
            if description in description_transfo.keys():
                description = description_transfo[description]
            elif "Nantissement" in description:
                description = "Pledge of securities"
            else:
                description = "Check PDF"

            priority = False
            for vip in priority_list:
                if (vip in exercer) | (vip.upper() in exercer):
                    priority = True
                    break           

            if (value > 50000) | (value == 0):
                transaction_date = texte_utilisable[texte_utilisable.find("DATE DE LA TRANSACTION : ") + len("DATE DE LA TRANSACTION : "): texte_utilisable.find("LIEU DE LA TRANSACTION :")].split()
                transaction_date = transaction_date[0] + "-" + transaction_date[1][:3] + "-" + transaction_date[2][-2:]
                reception_date = texte_utilisable[texte_utilisable.find("DATE DE RECEPTION DE LA NOTIFICATION : ") + len("DATE DE RECEPTION DE LA NOTIFICATION : "): texte_utilisable.find("COMMENTAIRES :")].split()
                reception_date = reception_date[0] + "-" + reception_date[1][:3] + "-" + reception_date[2][-2:]

                row_to_append = {
                    "REGISTRANT": exercer,
                    "COMPANY": company_name,
                    "OPERATION": description,
                    "QUANTITY": volume,
                    "AVG PRICE": round(price, 2),
                    "VALUE": value,
                    "TRADE DATE": transaction_date,
                    "Nature de la déclaration" : nature,
                    "DISCLOSING DATE": reception_date,
                    "Référence AMF" :référence_AMF,
                    "Commentaire AMF": commentaire_AMF     
                }
                df = df.append(row_to_append, ignore_index=True)
            else:
                pass

        df["QUANTITY"] = df["QUANTITY"].apply(lambda x: '{:,}'.format(int(x)).replace(',', ' '))
        df["AVG PRICE"] = df["AVG PRICE"].apply(lambda x: '{:,}'.format(x).replace(',', ' ')) + ["€" for elem in range(len(df["AVG PRICE"]))]
        df["VALUE"] = df["VALUE"].apply(lambda x: '{:,}'.format(int(round(x, 0))).replace(',', ' ')) + ["€" for elem in range(len(df["VALUE"]))]

        df = df.sort_values(by=["COMPANY", "OPERATION"])
        
        #Check if the excel file already exists and if not creat it
        if os.path.exists("Directos_Dealing _extract.xlsx"):
            pass
        else:
            df.to_excel("Directos_Dealing _extract.xlsx")
        

        excel_writer = StyleFrame.ExcelWriter(r"Directos_Dealing _extract.xlsx")
        sf = StyleFrame(df)

        sf.apply_style_by_indexes(indexes_to_style=sf.index,
                                cols_to_style=[elem for elem in sf.columns],
                                styler_obj=Styler(font_size=10, wrap_text=True))

        sf.apply_style_by_indexes(indexes_to_style=sf.index,
                                cols_to_style=["REGISTRANT"],
                                styler_obj=Styler(font_size=10))

        sf.apply_style_by_indexes(indexes_to_style=sf[sf["VALUE"].apply(lambda x: int(str(x)[:-1].replace(" ", ""))) > 25 * 10 ** 6],
                                cols_to_style=[elem for elem in sf.columns],
                                styler_obj=Styler( font_size=10))

        sf.apply_style_by_indexes(indexes_to_style=sf[sf["VALUE"].apply(lambda x: int(str(x)[:-1].replace(" ", ""))) > 25 * 10 ** 6],
                                cols_to_style=["REGISTRANT"],
                                styler_obj=Styler(font_size=10))

        sf.apply_style_by_indexes(indexes_to_style=sf[sf["OPERATION"] == "Pledge of securities"],
                                cols_to_style=[elem for elem in sf.columns],
                                styler_obj=Styler(font_size=10))

        sf.apply_style_by_indexes(indexes_to_style=sf[sf["OPERATION"] == "Pledge of securities"],
                                cols_to_style=[elem for elem in sf.columns],
                                styler_obj=Styler( font_size=10))

        sf.to_excel(
            excel_writer=excel_writer, 
            best_fit=[elem for elem in df.columns],
            startcol=1,
        )
        print(df)

        excel_writer.save()
        print("Report created")

  
    def get_DD(self, search_word , start_date , end_date) :
        #Function to extract (from the AMF website) all Directors's dealings between two dates and given a particular search word
        print("Launching...")
        start = datetime.now()
        os.chdir(self.path)
        files = glob.glob(self.path + "\\*")
        for f in files:
            os.remove(f) # Remove previous files
            
        lien = f"https://bdif.amf-france.org/fr?rechercheTexte={search_word}&&&dateDebut={start_date}&dateFin={end_date}&typesInformation=DD"
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : self.path}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
        chrome_options.add_argument("disable-infobars")
        chrome_options.add_argument("--start-maximized")
        chrome = webdriver.Chrome(self.path_driver,chrome_options=chrome_options)
        chrome.get(lien)
        nbr_download = 0 
        lien_folder_dd = self.path
        height = 0
        i=1
        while True:
            try :
                if chrome.find_element_by_xpath("/html/body/app-root/div/div/app-header/header/div[1]/div[2]/span").is_displayed():        
                    break
            except BaseException:
                t.sleep(0.5)
                
        if chrome.find_element_by_id("tarteaucitronPersonalize2").is_displayed():
            chrome.find_element_by_xpath("/html/body/div[2]/div[3]/button[1]").click()
        else :
            t.sleep(1.5)
            
        while True:
            try:   
                if chrome.find_element_by_xpath("/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[1]/h2").is_displayed():
                    Nombre_de_DD = chrome.find_element_by_xpath("/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[1]/h2").text
                    Nombre_de_DD = int(re.findall(r'\d+',Nombre_de_DD)[0])
                    print(f"Directors' Dealing number: {Nombre_de_DD}")
                    break
            except BaseException:
                t.sleep(0.5)
            
        if Nombre_de_DD == 0:
            print("No data for this period")
            chrome.quit()
            return

        while chrome.find_element_by_class_name("info").is_displayed(): 
            try :
                chrome.find_element_by_xpath(f"/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[2]/div[1]/ul/li[{i}]/div/app-result-list-view/a/mat-card/mat-card-actions/button").click()
                while len(os.listdir(lien_folder_dd)) == nbr_download:
                    t.sleep(0.01)
                t.sleep(1)
                nbr_download += 1
                i+=1
            except BaseException:
                try :
                    chrome.execute_script(f"window.scrollTo(0,{height})")
                    height = height + 72
                    t.sleep(0.3)
                    if chrome.find_element_by_xpath("/html/body/app-root/div/div/app-home-container/section/app-results-container/div[2]/div[2]/div/a").is_displayed():
                        chrome.find_element_by_xpath("/html/body/app-root/div/div/app-home-container/section/app-results-container/div[2]/div[2]/div/a").click()
                        t.sleep(0.3)
                    else:
                        continue
                except NoSuchElementException :
                    break    
        page_nb = 1
        for file in os.listdir(lien_folder_dd):
            os.rename(lien_folder_dd + "\\"+ file, lien_folder_dd + "\\"+ f"DD n°{page_nb}.pdf" )
            page_nb += 1
        print(f"EXEC TIME : {datetime.now() - start}, DD downloaded")
        chrome.quit()
        self.extract_DD()
        
    #Create the UX interface using Tkinter
        
    def UX(self):
        os.chdir(r"{}".format(self.package_directory))
        window=Tk()
        color= "#C42708"
        window.title('AMF advanced research')
        window.geometry("580x460")
        window.resizable(width=False, height=False)
        try:
            window.iconbitmap(r"amf_ico.ico")
        except BaseException :
            print("The path of the .ico image is probably wrong")
            pass
        window.config(background="#FFFFFF")

        style = ttk.Style()                     
        current_theme =style.theme_use()
        style.theme_settings(current_theme, {"TNotebook.Tab": 
                                            {"configure": {"padding": [130, 5]}}
                                            })  
        style.configure('TNotebook.Tab', font=('Calibri','10','bold'))
        style.configure("TNotebook.Tab", foreground="#003867") 

        style.configure('TFrame', background='white')

        #***********************************DD*****************************************************

        #Create the central logo
        
        width=200
        height = 200
        try :
            image = PhotoImage(file= r"AMF_logo.PNG").zoom(10).subsample(27)
            canvas= Canvas(window,width=width+6,height=height+100,bg="#FFFFFF",bd=0,highlightthickness=0)
            canvas.create_image( width/2, height/8, image=image)
            canvas.pack(expand=YES)
        except BaseException :
            print("The path of the logo is probably wrong")
            pass

        #Create the entry box
        keywords = Entry(window)
        keywords.place(x = 135, y = 150,height = 30,width = 305)
        keywords.config(borderwidth = 4, relief = "groove")
        
        #Create the entry box for the folder path
        lien_folder = Entry(window)
        lien_folder.place(x = 135,y = 370,height = 30,width = 305)
        lien_folder.config(borderwidth = 4, relief = "groove")

        #Add "Folder" text
        folder_1 = Label(window,text = "Link of the folder",bg = "white",font = ('Calibri', 14, 'bold'),fg ="#003867")
        folder_1.place(x=220,y=340)

        #Staring date
        start=Label(window,text = "Start date",bg = "white",font = ('Calibri', 14, 'bold'),fg = "#003867") 
        start.place(x = 157,y = 200)
        start_date = DateEntry(window, width = 10, year = 2022, month = 4, day = 1, 
        background ='darkblue', foreground ='white', borderwidth = 2,)
        start_date.place(x = 160,y = 230)

        #Ending date
        end=Label(window,text ="End date",bg = "white", font = ('Calibri', 14, 'bold'),fg = "#003867") 
        end.place(x = 333,y = 200)
        end_date = DateEntry(window, width = 10, year = 2022, month = 4, day = 1, 
        background = 'darkblue', foreground = 'white', borderwidth = 2)
        end_date.place(x = 332,y = 230)

        #Declaration type
        nom_page_1 = Label(window,text = "Directors' Dealings",bg = "white",font = ('Calibri', 14, 'bold'),fg = "#003867")
        nom_page_1.place(x = 205,y = 120)
        
        #Get the output excel path
        excel_output_path = lien_folder.get()
        
        def on_button():
            self.get_DD(keywords.get(),str(start_date.get_date()),str(end_date.get_date()))
            if lien_folder.get() != "":
                new_path =  r"{}{}Directos_Dealing _extract.xlsx".format(lien_folder.get(),"\\")
                new_path =  new_path.replace('"', "")
                shutil.move(r"Directos_Dealing _extract.xlsx", new_path)
            else :
                pass
    
        #Launch button
        btn = Button(window, text = 'Launch research',font = ('Calibri', 14, 'bold'),fg = "white",bg = "#003867",command = on_button)
        btn.place(x = 135,y = 280,height = 40,width = 305)
        menu = Menu(window)
        file_menu = Menu(menu, tearoff = 0)
        file_menu.add_command(label = "Exit", command = window.destroy)
        menu.add_cascade(label = "Options",menu = file_menu)
        window.config(menu = menu)
        window.mainloop()
    
if __name__ == "__main__":
    a = DD(r"C:\Users\ernes\OneDrive\Bureau\chromedriver.exe") #Replace with the path of your chromedriver
    a.UX()




