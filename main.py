import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from CSV_WRITER import  CSV_WRITER
import os
import datetime
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import math
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from Aktie import Aktie
import pandas
from Reader import Reader

def cookies():
    try:
        cookie_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="CybotCookiebotDialogBodyButtonDecline"]')))
        cookie_button.click()
    except:
        print("Cookie Pop Up nicht erschienen!")

def getdata():
    rd = Reader()
    rd.openExplorer()
    path = rd.getPath()
    data = pandas.read_csv(filepath_or_buffer=path, sep  = ";")
    return data["WKN"].to_list()


wkns = getdata()

liste = []

url = "https://account.stock3.com"

driver = webdriver.Chrome()

driver.get(url)

cookies()

time.sleep(0.5)

driver.find_element(By.ID, "input-username").send_keys("wild.marco@online.de")
time.sleep(0.6)
driver.find_element(By.ID, "input-password").send_keys("mawi#2021_rlph")
time.sleep(0.3)
driver.find_element(By.XPATH, '//*[@id="app"]/main/div[1]/form/div[3]/div/button').click()
time.sleep(2)

url = "https://terminal.stock3.com"
driver.get(url)

cookies()

driver.maximize_window()

driver.find_elements(By.CLASS_NAME, "widget__dropzone-button")[2].click()

time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[3]/div[2]/div[2]/div[2]/div[34]/div/div[2]/div[1]').click()

time.sleep(1)

driver.find_element(By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[3]/div[4]/div/div/div[2]/div[2]/div[2]/div').click()

time.sleep(2)

## ab hier Schleife über Liste und Liste mit dictionaries befuellen
for wert in wkns:
    if wert == "":
        continue
    for einzelaktie in liste:
        if wert in einzelaktie:
            continue

    search_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]')))

    search_button.click()

    time.sleep(1)

    actions = ActionChains(driver)

    actions.send_keys(wert)

    actions.send_keys(Keys.ENTER)

    actions.perform()

    time.sleep(0.5)

    peergroup_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[2]/div[2]/div[1]/div[3]')))
    peergroup_button.click()
    time.sleep(1)
    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")

    ## ab hier schleife über die Objekte in der Peergroup
    for i in range(1, 6):
        # print(soup.findAll("th"))
        daten = []
        driver.find_elements(By.CLASS_NAME, "button__label")[i+3].click()
        time.sleep(0.2)
        new_html = driver.page_source
        new_soup = BeautifulSoup(new_html, "html.parser")
        daten.append(new_soup.select(".accordion__parameter")[0].findAll("span")[0].text.strip())
        daten.append(soup.findAll("th")[i].text.strip())
        daten.append(new_soup.select(".accordion__parameter")[0].findAll("span")[1].text.strip())
        daten.append(new_soup.select_one(".industry").text.strip())
        daten.append(new_soup.select(".accordion__parameter")[6].select_one(".button").text.strip())
        # objekt = Aktie(soup.findAll("th")[i].text.strip(), wkn, isin, branche, sektor)
        driver.find_element(By.XPATH, '//*[@id="grid"]/div[4]/div[1]/div[1]/div[2]/i[6]').click()
        # print(soup.findAll("tr")[1].select_one("td").text)

        # Schleife über die Reihen in der Übersicht
        daten.append(soup.findAll("tr")[1].select("td")[i].select_one(".stock3Score__total").text.strip().replace("\u202f%", ""))
        for row in soup.findAll("tr")[2:]:
            daten.append(row.select("td")[i].text.strip().replace("\u202f%", ""))

        liste.append(daten)
        columns = ["WKN", "Name", "ISIN", "Branche", "Sektor", "Stock3Score", "Momentum & Vola", "Kursperform 6 M", "Kursperform 1 J", "Delta 52 Wochen Hoch", "Vola 1 J",
                   "Bewertung", "KUV", "Free Cash Flow", "PEG Ratio", "KGV (2023)", "KGV (2024)", "Kursziel", "Wachstum", "Umsatzwachstum ueber 5 J", "EPS-Wachstum ueber 5 J",
                   "Umsatzwachstum (2023)", "Umsatzwachstum (2024)", "Wachstum des verwaesserten Gewinns je Aktie (2023)", "Wachstum des verwaesserten Gewinns je Aktie (2024)",
                   "Qualitaet und Verschuldung", "Eigenkapitalrendite", "Eigenkapitalquote", "EBIT-Marge","Liquiditaet dritten Grades", "Liquiditaet zweiten Grades", "Verhaeltnis aus Schulden zum EK",
                   "Verhaeltnis aus Schulden und  Vermoegenswerten", "Zinsdeckungsgrad", "Dividende & Aktienrueckkaeufe", "Gesamtrendite", "Payout Ratio", "Dividende (2023)", "Dividende (2024"]

    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[2]/div[2]/div/div[2]'))).click()

    time.sleep(0.25)

df = pandas.DataFrame(liste, columns = columns)
folder = os.path.dirname(__file__)
filename = "stock3_" + datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%y_%H%M") + ".csv"
df.to_csv(os.path.join(folder, filename), sep=";", index = False)