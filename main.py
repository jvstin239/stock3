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
from webdriver_manager.chrome import ChromeDriverManager

def cookies():
    try:
        cookie_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="CybotCookiebotDialogBodyButtonDecline"]')))
        cookie_button.click()
    except:
        print("Cookie Pop Up nicht erschienen!")

# Funktion, die prüft, ob der Wert keine Zahl ist und ihn durch einen leeren String ersetzt
def replace_non_numeric(value):
    # Wenn der Wert keine Instanz von int oder float ist, ersetzen Sie ihn durch ''
    if not isinstance(value, (int, float)):
        return ''
    return value

def getdata():
    rd = Reader()
    rd.openExplorer()
    path = rd.getPath()
    data = pandas.read_csv(filepath_or_buffer=path, sep  = ";")
    return data["WKN"].to_list()


wkns = getdata()

liste = []

url = "https://account.stock3.com"

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get(url)

cookies()

time.sleep(0.5)

driver.find_element(By.ID, "input-username").send_keys("wild.marco@online.de")
time.sleep(0.6)
driver.find_element(By.ID, "input-password").send_keys("mawi#2021_rlph")
time.sleep(0.3)
driver.find_element(By.XPATH, '//*[@id="app"]/main/div[1]/form/div[3]/div/button').click()
time.sleep(2)

url = "https://terminal.stock3.com/#c/545737"
driver.get(url)

cookies()

driver.maximize_window()

time.sleep(20)


### ab hier wurde das Widget gesucht


#driver.get("https://terminal.stock3.com/#store/w")
# driver.find_elements(By.CLASS_NAME, "widget__dropzone-button")[2].click()

#WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='toolbox-item__title' and text()='stock3 Score']"))).click()
#driver.find_element(By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[3]/div[1]/div[2]/div[2]/div[32]/div/div[2]/div[1]').click()

#WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='button button--up' and text()='Widget hinzufügen']"))).click()

#driver.find_element(By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[3]/div[3]/div/div/div[2]/div[2]/div[2]/div').click()

#time.sleep(2)

wkns_bereits_drin = []

wkn_vorher = ""

## ab hier Schleife über Liste und Liste mit dictionaries befuellen
for wert in wkns:
    if wert == "":
        continue

##############################
    try:
    # search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]')))
        stock3_score_widget = driver.find_element(By.CSS_SELECTOR, 'div[data-w="score"]')

        # Locate the search button within this widget
        search_button = stock3_score_widget.find_element(By.CLASS_NAME, 'button--icon')
        search_button.click()

        time.sleep(1)

        actions = ActionChains(driver)

        actions.send_keys(wert)

        actions.perform()

        time.sleep(0.1)

        actions.send_keys(Keys.ENTER)

        actions.perform()

        time.sleep(0.5)

        first_html = driver.page_source
        first_soup = BeautifulSoup(first_html, "html.parser")

        stock3_check = first_soup.select_one(".stock3Score__total").text.strip()
        if stock3_check == "-":
            continue


        try:
            peergroup_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[2]/div[1]/div[2]/div[1]/div[1]/simple-button[2]')))
            peergroup_button.click()
        except Exception as error:
            print("Peergroup Button konnte nicht geklickt werden bei: " + wert)
            print(error)
            continue
        time.sleep(1)
        WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
        )
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        stock3_scores_check=[]
        for i in range(0, 5): stock3_scores_check.append(soup.select(".stock3Score__total")[i].text)

        if "-" in stock3_scores_check:
            print("WKN hat kein Score - " + str(wert))
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[2]/div[1]/div[2]/div[1]/div/simple-button[2]'))).click()
            continue
        table = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "tbl.ratingComparison"))
        )

        ## ab hier schleife über die Objekte in der Peergroup
        for i in range(1, 6):
            daten = []
            time.sleep(1)
            #time.sleep(5)
            driver.find_elements(By.CSS_SELECTOR, '.button__label.dr')[i-1].click()
            time.sleep(0.25)
            #driver.find_element(By.CLASS_NAME, 'switch__item').click()

            time.sleep(0.5)
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            new_html = driver.page_source
            new_soup = BeautifulSoup(new_html, "html.parser")

            # print(new_soup)


            #try:

            wkn = new_soup.select(".accordion__parameter-value")[0].findAll("span")[0].text.strip()
            # widget = new_soup.find('div', {'data-w': 'instrument1'})
            if wkn in wkns_bereits_drin:
                element_to_hover = driver.find_element(By.XPATH,
                                                       '//*[@id="grid"]/div[3]/div[1]/div[1]')

                # Perform the hover action
                actions = ActionChains(driver)
                actions.move_to_element(element_to_hover).perform()
                close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[1]/div[2]/div/i[6]')))
                close_icon.click()
                continue
            wkns_bereits_drin.append(wkn)
            #time.sleep(100000)
            #WKN
            daten.append(wkn)
            #Name
            daten.append(soup.findAll("th")[i].text.strip())
            #ISIN
            daten.append(new_soup.select(".accordion__parameter-value")[0].findAll("span")[1].text.strip())
            #Branche
            daten.append(new_soup.select_one(".industry").text.strip())
            #Sektor
            sektor = new_soup.select(".accordion__parameter-value")[6].select_one(".button").text.strip()
            daten.append(sektor)
            #daten.append(new_soup.findAll("Sektor")[0].select_one(".button").text.strip())
            # sell Kurs
            daten.append(new_soup.select(".box-item__price")[0].text.strip())
            # buy Kurs
            daten.append(new_soup.select(".box-item__price")[1].text.strip())
            # Land
            daten.append(new_soup.select_one(".country-items").text.strip())
            # anzahl Aktien
            daten.append(new_soup.select(".accordion__parameter-value")[8].text.strip().replace('\u202f', ' '))
            # Marktkapitalisierung
            daten.append(new_soup.select(".accordion__parameter-value")[9].text.strip().replace('\u202f', ' '))

            #except:
            #    print("Element hat kein Attribut text bei " + " in peergroup Position " )
            #    element_to_hover = driver.find_element(By.XPATH,
            #                                           '//*[@id="grid"]/div[7]')  # Replace with your locator



                # Perform the hover action
                #actions = ActionChains(driver)
                #actions.move_to_element(element_to_hover).perform()
                #driver.find_element(By.XPATH, '//*[@id="grid"]/div[4]/div[1]/div[1]/div[2]/i[6]').click()
                #close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '// *[ @ id = "grid"] / div[7] / div[1] / div[1] / div[2] / div / i[6]')))
                #close_icon.click()
                #continue

            element_to_hover = driver.find_element(By.XPATH,
                                                   '//*[@id="grid"]/div[3]/div[1]/div[1]')  # Replace with your locator

            # Perform the hover action
            actions = ActionChains(driver)
            actions.move_to_element(element_to_hover).perform()
            close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[1]/div[2]/div/i[6]')))
            close_icon.click()
            time.sleep(0.2)

            # Schleife über die Reihen in der Übersicht
            stock3_score = soup.findAll("tr")[1].select("td")[i].select_one(".stock3Score__total").text.strip().replace("\u202f%", "")
            # print("Stock 3 score " + wkn + ": " + str(stock3_score))

            if stock3_score == "-":
                # daten = []
                break
            else:
                daten.append(stock3_score)

            try:
                for row in soup.findAll("tr")[2:]:
                    zahl = row.select("td")[i].text.strip().replace("\u202f%", "")
                    if zahl not in ["neg.", "access denied", "instrument n/a", "-"]:
                        daten.append(zahl)
                    else:
                        daten.append("")
            except:
                print("Fehler in Schleife über die einzelnen Reihen bei " + " in Position " + str(i))
                continue
            liste.append(daten)
            #print(liste)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[2]/div[1]/div[2]/div[1]/div/simple-button[2]'))).click()

        # WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[2]/div[2]/div/div[2]'))).click()

        time.sleep(0.25)
#########################
    except Exception as error:
        print("Problem beim Abruf! Abbruch bei " + str(wert))
        print("error: ", error)
        break

columns = ["WKN",
            "Name",
            "ISIN",
            "Branche",
            "Sektor",
            "Sell-Kurs",
            "Buy-Kurs",
            "Land",
            "Anzahl_Aktien",
            "Marktkapitalisierung",
            "Stock3Score",
            "Momentum_&_Vola",
            "Vola_1_J",
            "Delta_52_Wochen_Hoch",
            "Kursperform_6_M",
            "Kursperform_1_J",
            "Bewertung",
            "Kursziel",
            "KGV_(Jahr)",
            "KGV_(Jahr_+1)",
            "Free_Cash_Flow",
            "KUV",
            "PEG_Ratio",
            "Wachstum",
            "Umsatzwachstum_(Jahr)",
            "Umsatzwachstum_(Jahr_+1)",
            "Wachstum_des_verwaesserten_Gewinns_je_Aktie_(Jahr)",
            "Wachstum_des_verwaesserten_Gewinns_je_Aktie_(Jahr_+1)",
            "Umsatzwachstum_ueber_5_J",
            "EPS_Wachstum_ueber_5_J",
            "Qualitaet_und_Verschuldung",
            "Eigenkapitalrendite",
            "Eigenkapitalquote",
            "EBIT_Marge",
            "Liquiditaet_zweiten_Grades",
            "Liquiditaet_dritten_Grades",
            "Verhaeltnis_aus_Schulden_und_Vermoegenswerten",
            "Verhaeltnis_aus_Schulden_zum_EK",
            "Zinsdeckungsgrad",
            "Dividende",
            "Payout_Ratio",
            "Dividende_(Jahr)",
            "Dividende_(Jahr_+1)",
            "Dividende_(Jahr_+2)",
            "Dividende_(Jahr_+3)",
            "fuenfjaehrige_Wachstumsrate_Dividende_je_Aktie"]

df = pandas.DataFrame(liste, columns = columns)
# df['Zinsdeckungsgrad'] = df['Zinsdeckungsgrad'].apply(replace_non_numeric)
#older = os.path.dirname(__file__)
filename = "stock3_" + datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%y_%H%M") + ".csv"
#df.to_csv(os.path.join(folder, filename), sep=";", index = False, encoding = "utf-8")
df.to_csv("//Master/F/User/Microsoft Excel/Privat/Börse/Stock3_Bewertungen/" + filename, sep=";", index = False, encoding = "utf-8")