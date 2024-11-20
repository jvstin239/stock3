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

url = "https://terminal.stock3.com"
driver.get(url)

cookies()

driver.maximize_window()

time.sleep(1)

driver.get("https://terminal.stock3.com/#store/w")
# driver.find_elements(By.CLASS_NAME, "widget__dropzone-button")[2].click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='toolbox-item__title' and text()='stock3 Score']"))).click()
#driver.find_element(By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[3]/div[1]/div[2]/div[2]/div[32]/div/div[2]/div[1]').click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='button button--up' and text()='Widget hinzufügen']"))).click()

#driver.find_element(By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[3]/div[3]/div/div/div[2]/div[2]/div[2]/div').click()

time.sleep(2)

wkns_bereits_drin = []

wkn_vorher = ""

## ab hier Schleife über Liste und Liste mit dictionaries befuellen
for wert in wkns:
    if wert == "":
        continue

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

        try:
            peergroup_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='widget__col widget__col--left']/div[text()='Mit Peergroup vergleichen']")))
            peergroup_button.click()
        except:
            print("Peergroup Button konnte nicht geklickt werden bei: " + wert)
            continue
        time.sleep(1)
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        table = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "tbl.ratingComparison"))
        )

        ## ab hier schleife über die Objekte in der Peergroup
        for i in range(1, 6):
            daten = []
            time.sleep(1)
            driver.find_elements(By.CSS_SELECTOR, '.button__label.dr')[i-1].click()

            time.sleep(0.5)
            new_html = driver.page_source
            new_soup = BeautifulSoup(new_html, "html.parser")

            # print(new_soup)


            try:

                wkn = new_soup.select(".accordion__parameter")[10].findAll("span")[0].text.strip()
                widget = new_soup.find('div', {'data-w': 'instrument1'})
                # wkn = soup.select_one(".accordion__parameter-value span").text
                #print(wkn)
                if wkn in wkns_bereits_drin:
                    element_to_hover = driver.find_element(By.XPATH,
                                                           '//*[@id="grid"]/div[7]')  # Replace with your locator

                    # Perform the hover action
                    actions = ActionChains(driver)
                    actions.move_to_element(element_to_hover).perform()
                    close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                        (By.XPATH, '// *[ @ id = "grid"] / div[7] / div[1] / div[1] / div[2] / div / i[6]')))
                    close_icon.click()
                    continue
                wkns_bereits_drin.append(wkn)
                #time.sleep(100000)
                #WKN
                daten.append(wkn)
                #Name
                daten.append(soup.findAll("th")[i].text.strip())
                #ISIN
                daten.append(new_soup.select(".accordion__parameter")[10].findAll("span")[1].text.strip())
                #Branche
                daten.append(widget.select_one(".industry").text.strip())
                #Sektor
                sektor = widget.select(".accordion__parameter")[6].select_one(".button").text.strip()
                daten.append(sektor)
                #daten.append(new_soup.findAll("Sektor")[0].select_one(".button").text.strip())
                # sell Kurs
                daten.append(widget.select(".box-item__price")[0].text.strip())
                # buy Kurs
                daten.append(widget.select(".box-item__price")[1].text.strip())
                # Land
                daten.append(widget.select_one(".country-items").text.strip())
                # anzahl Aktien
                daten.append(widget.select(".accordion__parameter-value")[8].text.strip())
                # Marktkapitalisierung
                daten.append(widget.select(".accordion__parameter-value")[9].text.strip())
            except:
                print("Element hat kein Attribut text bei " + " in peergroup Position " )
                element_to_hover = driver.find_element(By.XPATH,
                                                       '//*[@id="grid"]/div[7]')  # Replace with your locator

                # Perform the hover action
                actions = ActionChains(driver)
                actions.move_to_element(element_to_hover).perform()
                driver.find_element(By.XPATH, '//*[@id="grid"]/div[4]/div[1]/div[1]/div[2]/i[6]').click()
                close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '// *[ @ id = "grid"] / div[7] / div[1] / div[1] / div[2] / div / i[6]')))
                close_icon.click()
                continue

            element_to_hover = driver.find_element(By.XPATH,
                                                   '//*[@id="grid"]/div[7]')  # Replace with your locator

            # Perform the hover action
            actions = ActionChains(driver)
            actions.move_to_element(element_to_hover).perform()
            close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '// *[ @ id = "grid"] / div[7] / div[1] / div[1] / div[2] / div / i[6]')))
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

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[6]/div[1]/div[2]/div[2]/div/div[2]'))).click()

        # WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/div[1]/div[2]/div[2]/div/div[2]'))).click()

        time.sleep(0.25)

    except Exception as error:
        print("Problem beim Abruf! Abbruch bei " + str(wert))
        print("error: ", error)
        break

columns = ["WKN", "Name", "ISIN", "Branche", "Sektor", "Sell-Kurs", "Buy-Kurs", "Land", "Anzahl_Aktien", "Marktkapitalisierung", "Stock3Score", "Momentum_&_Vola",
                   "Kursperform_6_M", "Kursperform_1_J", "Delta_52_Wochen_Hoch", "Vola_1_J",
                   "Bewertung", "KUV", "Free_Cash_Flow", "PEG_Ratio", "KGV_(Jahr)", "KGV_(Jahr_+1)", "Kursziel", "Wachstum", "Umsatzwachstum_ueber_5_J", "EPS_Wachstum_ueber_5_J",
                   "Umsatzwachstum_(Jahr)", "Umsatzwachstum_(Jahr_+1)", "Wachstum_des_verwaesserten_Gewinns_je_Aktie_(Jahr)", "Wachstum_des_verwaesserten_Gewinns_je_Aktie_(Jahr_+1)",
                   "Qualitaet_und_Verschuldung", "Eigenkapitalrendite", "Eigenkapitalquote", "EBIT_Marge","Liquiditaet_dritten_Grades", "Liquiditaet_zweiten_Grades", "Verhaeltnis_aus_Schulden_zum_EK",
                   "Verhaeltnis_aus_Schulden_und_Vermoegenswerten", "Zinsdeckungsgrad", "Dividende_&_Aktienrueckkaeufe", "Gesamtrendite", "Payout_Ratio", "Dividende_(Jahr)", "Dividende_(Jahr_+1)"]

df = pandas.DataFrame(liste, columns = columns)
# df['Zinsdeckungsgrad'] = df['Zinsdeckungsgrad'].apply(replace_non_numeric)
folder = os.path.dirname(__file__)
filename = "stock3_" + datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%y_%H%M") + ".csv"
df.to_csv(os.path.join(folder, filename), sep=";", index = False, encoding = "utf-8")
#df.to_csv("//Master/F/User/Microsoft Excel/Privat/Börse/Stock3_Bewertungen/" + filename, sep=";", index = False, encoding = "utf-8")