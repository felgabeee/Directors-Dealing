#!/usr/bin/env python
# coding: utf-8

# In[48]:


pip freeze > requirements.txt


# In[ ]:


pip install requirements.txt


# In[25]:


import pandas as pd
import numpy as np

import requests
from bs4 import BeautifulSoup
import io
import tia.bbg.datamgr as dm

from datetime import timedelta
from datetime import time
from datetime import date
from datetime import datetime
import time as t
import locale
from datetime import date

import os
import glob
from openpyxl import load_workbook
from styleframe import StyleFrame
from styleframe import Styler

import re
import PyPDF2
from fuzzywuzzy import process
import unidecode
import warnings

import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
import shutil


from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import StaleElementReferenceException

warnings.filterwarnings("ignore")

    
def get_dd_online(path_driver):
    print("Launching...")
    language = locale.getdefaultlocale()[0]
    FR_desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Bureau')
    EN_desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    if language == "fr_FR":
        newpath = FR_desktop_path +"\\FR_dd"
        desktop_path=FR_desktop_path
        if not os.path.exists(newpath):
            os.makedirs(newpath)
            path=newpath
        else :
            path=newpath
    else:
        newpath = EN_desktop_path +"\\FR_dd" 
        if not os.path.exists(newpath):
            os.makedirs(newpath)
            path=newpath
        else :
            path=newpath
            desktop_path=EN_desktop_path
            
    debut = datetime.now()
    today= datetime.today() 
    today = today.strftime("%d/%m/%Y")[0:2] 
    mtn = str(datetime.today())
    ajd = mtn[:mtn.find(" ")]
    jour = int(mtn[8:10])
    mois = mtn[5:7]
    an = mtn[:4]
    files = glob.glob(f'{path}\\*')
    for f in files:
        os.remove(f)  
    if datetime.today().weekday() == 0:
        jour -= 3
    else:
        jour -= 1
    if len(str(jour)) == 1:
        jour = "0" + str(jour)
    lien = f"https://bdif.amf-france.org/fr?dateDebut={an}-{mois}-{jour}&dateFin={an}-{mois}-{jour}&typesInformation=DD"
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : path}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
    chrome_options.add_argument("disable-infobars")
    chrome_options.add_argument("--start-maximized")
    chrome = webdriver.Chrome(path_driver,chrome_options=chrome_options)
    chrome.get(lien)
    nbr_download =0 
    lien_folder_dd = path
    height=0
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
                Nombre_de_DD=chrome.find_element_by_xpath("/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[1]/h2").text
                Nombre_de_DD=int(re.findall(r'\d+',Nombre_de_DD)[0])
                print(f"Directors' Dealing number: {Nombre_de_DD}")
                break
        except BaseException:
            t.sleep(0.5)
        
    if Nombre_de_DD ==0:
        print("Pas de déclarations pour cette période")
        chrome.quit()
        return

    while chrome.find_element_by_class_name("info").is_displayed(): 
        try :
            chrome.find_element_by_xpath(f"/html/body/app-root/div/div/main/app-home-container/section/app-results-container/div[2]/div[1]/ul/li[{i}]/div/app-result-list-view/a/mat-card/mat-card-actions/button").click()
            while len(os.listdir(lien_folder_dd)) ==nbr_download:
                t.sleep(0.01)
            t.sleep(1)
            nbr_download+=1
            i+=1
        except BaseException:
            try :
                chrome.execute_script(f"window.scrollTo(0,{height})")
                height= height + 72
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
        page_nb+=1
    print(f"EXEC TIME : {datetime.now() - debut}, les DD ont bien été enregistrés")
    chrome.quit()
    

    

