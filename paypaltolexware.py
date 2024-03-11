from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementClickInterceptedException
from datetime import datetime, timedelta
from enum import Enum
import weasyprint
from weasyprint import HTML
import time
import locale
import re
import os
import csv
import pandas as pd
import win32com.client
import pdfkit
import autoit

locale.setlocale(locale.LC_TIME, "de_DE")

# Specify the debugging address for the already opened Chrome browser
debugger_address = 'localhost:8989'
optionsDEBUG = webdriver.ChromeOptions()
optionsDEBUG.add_experimental_option("debuggerAddress", "localhost:8989")

# Create a new instance of the Chrome driver
# optionsPaypal = webdriver.ChromeOptions()
# optionsPaypal.add_experimental_option("detach", True)
# driverPaypal = webdriver.Chrome(options=optionsDEBUG)

# optionsEbay = webdriver.ChromeOptions()
# optionsEbay.add_experimental_option("detach", True)
# driverEbay = webdriver.Chrome(options=optionsEbay)

optionsAmazon = webdriver.ChromeOptions()
optionsAmazon.add_experimental_option("detach", True)
driverAmazon = webdriver.Chrome(options=optionsAmazon)

optionsLexware = webdriver.ChromeOptions()
optionsLexware.add_experimental_option("detach", True)
driverLexware = webdriver.Chrome(options=optionsLexware)
waitLexware = WebDriverWait(driverLexware, 10)

# optionsStrato = webdriver.ChromeOptions()
# optionsStrato.add_experimental_option("detach", True)
# driverStrato = webdriver.Chrome(options=optionsStrato)
# waitStrato = WebDriverWait(driverStrato, 300)
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Iterate through all accounts to find the specified account, then get the folder
account = next(acc for acc in outlook.Folders if acc.Name == 'ebay@yourcamera.de')
folder = account.Folders['Posteingang']

messages = [message for message in folder.Items if message.Subject.startswith("Bestellung bestätigt:")]

class Type(Enum):
    ZahlungGesendetAn           = 1                #https://www.paypal.com/activity/payment/4SU601425B100803B  keine mitteilung->https://www.paypal.com/activity/payment/2GU03187EX4103709?Z3JncnB0=
    ZahlungGesendetEbay         = 2
    GeldEingezahltVon           = 3
    RückzahlungErfolgtAn        = 4                #https://www.paypal.com/activity/payment/4XX081207P2847308
    ZahlungIstEingegangenVon    = 5
    RechnungErhalten            = 6

    BezahlungDerBestellung      = 7
    Erstattung                  = 8

    NochUnbekannt               = 9

class Position:
    def __init__(self, index, datum, name, produkt, brutto, netto, typ, transactioncode):
        self.index = index
        self.datum = str(datum)
        self.name = name
        self.produkt = produkt
        self.brutto = brutto
        self.netto = netto
        self.typ = typ
        self.transactioncode = transactioncode
        self.bearbeitet = False

    def bearbeitung():
        raise NotImplementedError("Subclass must implement abstract method")
    
    def search_lexware(self):

        print("LEXWARE SEARCHING FOR: " + self.name + "\n")
        driverLexware.get("https://app.lexoffice.de/vouchers#!/VoucherList/?filter=lastedited")
        time.sleep(3)
        driverLexware.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/section/div[5]/input').send_keys(str(self.name))
                                                            
        rechnungenlist = driverLexware.find_elements(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[3]/div')

        self_date = datetime.strptime(self.datum, "%d.%m.%Y")
        start_date = self_date - timedelta(days=31)
        end_date = self_date + timedelta(days=31)

        r = 1
        for rechnung in rechnungenlist:
            try:
                time.sleep(2)
                rechnungprodukt = driverLexware.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[3]/div[' + str(r) + ']/div/div[1]/div[2]/div[1]/span[3]').text
                _rechnungbetrag = driverLexware.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[3]/div[' + str(r) + ']/div/div[2]/div/div[1]').text
                rechnungbetrag = float(_rechnungbetrag.replace('.', '').replace(',', '.').replace(' €', ''))
                self.transactioncode = driverLexware.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[3]/div[' + str(r) + ']/div/div[1]/div[2]/div[2]/span/span[2]').text
                _rechnungdatum = driverLexware.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[3]/div[' + str(r) + ']/div/div[1]/div[2]/div[2]/span').text
                _rechnungdatum = re.search(r'\d{2}\.\d{2}\.\d{4}', _rechnungdatum)
                rechnungdatum = datetime.strptime(_rechnungdatum.group(0), '%d.%m.%Y')

            except:
                #Keine Rechnungen Vorhanden?
                if (driverLexware.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[2]').text == "Keine Belege vorhanden"):
                    input('RECHNUNG NICHT GEFUNDEN, BITTE SCHREIBEN FÜR: '+ self.name)
                pass
                return 1
            
            if self.typ == Type.Erstattung:
                GSorRG = 'GS'
            else:
                GSorRG = '0'

            if (abs(rechnungbetrag - self.brutto) <= 100 or abs(rechnungbetrag - self.netto) <= 100) and start_date <= rechnungdatum <= end_date and self.transactioncode.startswith(str(GSorRG)) :
                print("RechnungGutschrift gefunden\n")
                self.produkt = rechnungprodukt

                driverLexware.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div/div[1]/section/div/div[2]/div[3]/div[' + str(r) + ']/div/div[2]/div/div[1]').click()
                original_window = driverLexware.current_window_handle
                time.sleep(1)   
                driverLexware.find_element(By.XPATH, '//*[@id="grldActionBar"]/div/div[3]/a/i').click()
                waitLexware.until(EC.number_of_windows_to_be(2))
                
                #switch Tab
                for window_handle in driverLexware.window_handles:
                    if window_handle != original_window:
                        driverLexware.switch_to.window(window_handle)
                        break
                
                time.sleep(1)
                autoit.win_activate(str(driverLexware.title))
                time.sleep(1)
                autoit.send("^p")
                time.sleep(2)
                autoit.send("{ENTER}")
                time.sleep(1)

                autoit.send("^w")
                driverLexware.switch_to.window(original_window)
                autoit.win_activate(str(driverLexware.title))
               
            r = r + 1
        return 0
    #TODO
    def Lexware_rechnung_schreiben():
        print("LEXWARE RECHNUNG SCHREIBEN\n")
        input("Enter any Value to continue\n")
        input("Are your sure?\n")  

class PaypalPosition (Position):
    def __init__(self, index, positiondate, positionbrutto, positionnetto, transactioncode):
        super().__init__(index, positiondate, '', '', positionbrutto, positionnetto, Type.NochUnbekannt, transactioncode)
        self.mitteilung = ""
        self.ebayname = ""

    def bearbeitung(self):

        driverPaypal.get("https://www.paypal.com/activity/payment/"+ self.transactioncode)

        TransactionHeader= driverPaypal.find_element(By.CSS_SELECTOR,'section.TDHeader[data-testid="tdheader_section"]').text
        print("BEARBEITE:\n" + TransactionHeader + "\n")


        while not self.bearbeitet:
            

            #Geld eingezahlt von     #Nur Drucken   brutto,            
            if re.match(r'^Geld eingezahlt von', TransactionHeader, re.IGNORECASE):

                self.PaypalType = Type.GeldEingezahltVon
                print("DRUCKE SEITE AUS: " + "https://www.paypal.com/activity/payment/" + self.transactioncode + "\n")
                autoit.win_activate(str(driverPaypal.title))
                autoit.send("^p")
                time.sleep(1)
                autoit.send("{ENTER}")
                time.sleep(1)
                self.bearbeitet = True

            #Zahlung gesendet an eBay S.a.r.l.  #Mitteilunf von eBay S.a.r.l.  Kein Produkt -> EBAY Transactionsnummer!
            elif re.match(r'^Zahlung gesendet an eBay S.a.r.l.', TransactionHeader, re.IGNORECASE):

                self.PaypalType = Type.ZahlungGesendetEbay
                mitteilung = driverPaypal.find_element(By.CSS_SELECTOR,'body > div:nth-child(1) > div.td-sections > div > div > div.details > section.Notes.pagebreak > div > p').text
                print (mitteilung)
                self.mitteilung = re.search(r'\b\d{2}-\d{5}-\d{5}\b', mitteilung).group(0)
                
                self.search_emails('ebay@yourcamera.de', 'Posteingang')
                self.bearbeitet = True
            
            # https://www.paypal.com/activity/payment/4SU601425B100803B
            # name, butto, Mitteilung an Name = Panasonic HC X1000
            elif re.search(r'Zahlung gesendet an (.*?)\n', TransactionHeader, re.IGNORECASE):

                self.PaypalType = Type.ZahlungGesendetAn                                        #TODO group 1 or 0???
                self.name = re.search(r'Zahlung gesendet an (.*?)\n', TransactionHeader, re.IGNORECASE).group(1)
                try:    
                    self.mitteilung = driverPaypal.find_element(By.CSS_SELECTOR,'p.col-sm-8.contentAlignedWithLabel').text
                    if self.mitteilung.replace('-', '').isdigit():
                        self.search_emails('ebay@yourcamera.de', 'Posteingang')
                except NoSuchElementException:
                    print("Keine Mitteilung!!!\n")
                input("KAUFBELEG??? " + str(self.name) + "\n")
                self.bearbeitet = True

            #Mitteilung von Kunde and Christian zb. ;Re nr 022412; https://www.paypal.com/activity/payment/00U07052TN0737336
            elif re.search(r'Zahlung ist eingegangen von ', TransactionHeader, re.IGNORECASE):
                self.PaypalType = Type.BezahlungDerBestellung
                self.name = re.search(r'Zahlung ist eingegangen von (.*?)\n', TransactionHeader, re.IGNORECASE).group(1)
                self.search_lexware()
                self.bearbeitet = True
            
            #
            elif re.match(r'^Rechnung erhalten', TransactionHeader, re.IGNORECASE):
                self.PaypalType = Type.RechnungErhalten
                input("RECHNUNG????\n")

            #Gutschrift mit Namen finden
            elif re.search(r'Rückzahlung erfolgt an (.*?)\n', TransactionHeader, re.IGNORECASE):
                self.PaypalType = Type.Erstattung
                self.name = re.search(r'Rückzahlung erfolgt an (.*?)\n', TransactionHeader, re.IGNORECASE).group(1)
                self.search_lexware()
                self.bearbeitet = True



        #TODO
        #Rechnung erhalten keine info     https://www.paypal.com/activity/payment/5U646543HK876761F
        #Geld Abgebucht auf (Credit/Bankkonto)
            
        print("BEARBEITET: \n" + str(self.bearbeitet))
        print("Datum: " + self.datum + "\n")
        print("Name: " + self.name + "\n")
        print("Typ: " + str(self.typ) + "\n")
        print("Produkt: " + self.produkt + "\n")
        print("Brutto: " + str(self.brutto) + "\n")
        print("Netto: " + str(self.netto) + "\n")
        print("Tran: " + self.transactioncode + "\n")

        
        # U P D A T E   C S V
        df_paypal.at[self.index, 'Datum'] = self.datum
        df_paypal.at[self.index, 'Name'] = self.name
        df_paypal.at[self.index, 'Ebayname'] = self.ebayname
        df_paypal.at[self.index, 'Produkt'] = self.produkt
        df_paypal.at[self.index, 'Transactioncode'] = self.transactioncode
        df_paypal.at[self.index, 'Brutto'] = self.brutto
        df_paypal.at[self.index, 'Netto'] = self.netto
        df_paypal.at[self.index, 'Mitteilung'] = self.mitteilung
        df_paypal.at[self.index, 'Type'] = self.typ
        df_paypal.at[self.index, 'Bearbeitet'] = self.bearbeitet
        df_paypal.to_csv('Paypal.csv', index=False, sep=';')

        self.bearbeitet = True

    def search_emails(self, account_name, folder_name):

        search_string = self.mitteilung
        for message in messages:

            if not isinstance(search_string, str):
                search_string = str(search_string)

            if search_string in message.Body:
                print(f"PRINTING: {message.Subject}\n")

                # Email as HTML
                with open('email.html', 'w', encoding='utf-8') as file:
                    file.write(message.HTMLBody)
                
                # HTML as PDF
                weasyprint.HTML('email.html').write_pdf("email.pdf")

                # P R I N T I N G   P D F
                os.startfile("email.pdf", "print")          

class AmazonPosition (Position):
    def __init__(self, index, positiondate, positionprodukt, positionnetto, positiontyp, transactioncode):
        super().__init__(index, positiondate, "", positionprodukt, 0, positionnetto, positiontyp, transactioncode)
        self.mitteilung = ""

    def bearbeitung(self):
        driverAmazon.get('https://sellercentral.amazon.de/orders-v3/order/' + self.transactioncode)
        time.sleep(2)
        buyerinfo = driverAmazon.find_elements(By.CSS_SELECTOR, 'div[data-test-id="shipping-section-buyer-address"]')
        buyerinfo = buyerinfo[0].text
        buyerinfo = buyerinfo.strip().split("\n")
        self.name = buyerinfo[0]
        #address = buyerinfo[1]
        #zipcode, city, region = buyerinfo[2].split(" ")[0], buyerinfo[2].split(",")[1],buyerinfo[2].split(",")[2].strip()
        try:
            self.produkt = driverAmazon.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div').text
        except:
            try:
                self.produkt = driverAmazon.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div').text
            except:
                pass

        try:
            self.brutto = float(driverAmazon.find_element(By.XPATH, '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[7]/div/table[1]/tbody/div[3]/div[2]/span').text.replace('.', '').replace(',', '.').replace('€','').replace('$','').replace('-',''))
        except:
            try:
                self.brutto = float(driverAmazon.find_element(By.XPATH, '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[7]/div/table[1]/tbody/div[3]/div[2]/span').text.replace('.', '').replace(',', '.').replace('€','').replace('$','').replace('-',''))
            except:
                pass
            
        self.search_lexware()
        
        print("BEARBEITET: \n" + str(self.bearbeitet))
        print("Datum: " + self.datum + "\n")
        print("Name: " + self.name + "\n")
        print("Typ: " + self.typ + "\n")
        print("Produkt: " + self.produkt + "\n")
        print("Brutto: " + str(self.brutto) + "\n")
        print("Netto: " + str(self.netto) + "\n")
        print("Tran: " + self.transactioncode + "\n")

        
        # U P D A T E   C S V
        df_amazon.at[self.index, 'Datum'] = self.datum
        df_amazon.at[self.index, 'Name'] = self.name
        df_amazon.at[self.index, 'Produkt'] = self.produkt
        df_amazon.at[self.index, 'Transactioncode'] = self.transactioncode
        df_amazon.at[self.index, 'Brutto'] = self.brutto
        df_amazon.at[self.index, 'Netto'] = self.netto
        df_amazon.at[self.index, 'Mitteilung'] = self.mitteilung
        df_amazon.at[self.index, 'Type'] = self.typ
        df_amazon.at[self.index, 'Bearbeitet'] = self.bearbeitet
        df_amazon.to_csv('Amazon.csv', index=False, sep=';')

        self.bearbeitet = True 

class EbayPosition (Position):
    def __init__(self, index, positiondate, name, produkt, positionbrutto, positionnetto, positiontyp, transactioncode):
        if str(positiontyp).startswith('Bestellung'):
            positiontyp = Type.BezahlungDerBestellung
        if str(positiontyp).startswith('Rückerstattung'):
            positiontyp = Type.Erstattung
        
        super().__init__(index, positiondate, name, produkt, positionbrutto, positionnetto, positiontyp, transactioncode)
        self.mitteilung = ""
        
    def bearbeitung(self):

        if self.search_lexware():
            print()
            #driverEbay.get('https://www.ebay.de/mesh/ord/details?orderid=' + str(self.transactioncode))
            #print page

        print("BEARBEITET: \n" + str(self.bearbeitet))
        print("Datum: " + self.datum + "\n")
        print("Name: " + self.name + "\n")
        print("Typ: " + str(self.typ) + "\n")
        print("Produkt: " + self.produkt + "\n")
        print("Brutto: " + str(self.brutto) + "\n")
        print("Netto: " + str(self.netto) + "\n")
        print("Tran: " + self.transactioncode + "\n")

        
        # U P D A T E   C S V
        df_ebay.at[self.index, 'Datum'] = self.datum
        df_ebay.at[self.index, 'Name'] = self.name
        df_ebay.at[self.index, 'Produkt'] = self.produkt
        df_ebay.at[self.index, 'Transactioncode'] = self.transactioncode
        df_ebay.at[self.index, 'Brutto'] = self.brutto
        df_ebay.at[self.index, 'Netto'] = self.netto
        df_ebay.at[self.index, 'Mitteilung'] = self.mitteilung
        df_ebay.at[self.index, 'Type'] = self.typ
        df_ebay.at[self.index, 'Bearbeitet'] = self.bearbeitet
        df_ebay.to_csv('Ebay.csv', index=False, sep=';')

        self.bearbeitet = True
            


def expand_shadow_element(element, driver):

    shadow_root = driver.execute_script('return arguments[0].shadowRoot', element)
    return shadow_root

def german_month_to_number(month_name):
    months = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez']
    return months.index(month_name) + 1

def get_all_open_tabs(driver):

    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        print(driver.current_url)

def create_df_csv(file_path):

    if not os.path.isfile(file_path):
        
        df = pd.DataFrame(columns=['Index', 'Datum', 'Name', 'Ebayname', 'Produkt', 'Mitteilung', 'Brutto', 'Netto', 'Transactioncode', 'Type', 'Bearbeitet'])
        df.to_csv(file_path, index=False, sep=';')

    else:
        df = pd.read_csv(file_path, sep=';', true_values=['True'], false_values=['False'])

    return df

def start_everything():
    
    # L E X W A R E
    driverLexware.get("https://app.lexoffice.de/vouchers#!/VoucherList/?filter=lastedited")
    try:
        driverLexware.find_element(By.CSS_SELECTOR,'[name="username"]').send_keys("info@yourcamera.de")
        driverLexware.find_element(By.CSS_SELECTOR,'[name="password"]').send_keys("Aachen12!")
        driverLexware.find_element(By.CSS_SELECTOR,'#mui-3').click()
    except:
        pass
    

    # P A Y P A L
    #driverPaypal.get("https://www.paypal.com/reports/financialSummaries/financialSummary")
  

    # E B A Y
    #driverEbay.get('https://www.ebay.de/mes/transactionlist?sh=true')
    # try:
    #     driverEbay.find_element(By.XPATH, '//*[@id="userid"]').send_keys('ebay@yourcamera.de')
    #     driverEbay.find_element(By.XPATH, '//*[@id="signin-continue-btn"]').click()
    #     time.sleep(1)
    #     driverEbay.find_element(By.XPATH, 'ebay@yourcamera.de').send_keys('Lizenzero12')
    #     driverEbay.find_element(By.XPATH, '//*[@id="sgnBt"]').click()
    # except:
    #     pass


    # A M A Z O N
    driverAmazon.get("https://sellercentral.amazon.de/payments/event/view?resultsPerPage=10&pageNumber=1")
    try:
        driverAmazon.find_element(By.XPATH, '//*[@id="ap_email"]').send_keys("televerkauf99@web.de")
        driverAmazon.find_element(By.XPATH, '//*[@id="ap_password"]').send_keys("aachen12")
        driverAmazon.find_element(By.XPATH, '//*[@id="signInSubmit"]').click()
    except:
        pass

    input("Enter any Value to continue\n")
    print("Continue\n")    
    input("Are your sure?\n")
    print("Continue\n\n")



def fill_df_paypal(df_paypal):

      #BERICHT ERSTELLEN//ALLE TRANSAKTIONEN ANZEIGEN
    try:
        driverPaypal.find_element(By.CSS_SELECTOR,'button.ppvx_btn___5-12-9.btn-btn-primary-csrSubmit[data-testid="FSRCreateReport"]').click()
        time.sleep(3)
        driverPaypal.find_element(By.CSS_SELECTOR,'button.linkButton[data-testid="linkButton"]').click()
        time.sleep(1)
    except NoSuchElementException:
        pass

    
    #SHOW MORE BUTTON
    i = 0
    while i == 0:
        time.sleep(1)
        try:
            driverPaypal.find_element(By.CSS_SELECTOR,'button.ppvx_btn___5-12-9.ppvx_btn--secondary___5-12-9#Showmore[data-testid="showMoreBtn"]').click()
        except NoSuchElementException:
            i = 1
            pass
        except StaleElementReferenceException:
            i = 1
            pass
        # except ElementClickInterceptedException:
        #     i = 1
        #     pass

    alltransactionslist = driverPaypal.find_elements(By.CSS_SELECTOR,'tr[data-testid^="tableRow"]')                 #List of all Elements that contain "tabelRow" as suffix
    alltransactionstable = driverPaypal.find_element(By.CSS_SELECTOR,'div.FSRListTable')                            #One single Element to narrow down search(under this element) 

    rowcounter = 0
    positionIndex = 1
    for transaction in alltransactionslist:
        
        t = alltransactionstable.find_element(By.CSS_SELECTOR,'tr[data-testid="tableRow' + str(rowcounter) + '"]')
        cells = t.find_elements(By.CSS_SELECTOR,'td')

        brutto = float(cells[4].text.replace('.', '').replace(',', '.'))
        netto = float(cells[5].text.replace('.', '').replace(',', '.'))
        rowcounter = rowcounter + 1
        
        if brutto > 50 or brutto < -50:
            if not df_paypal[df_paypal['Transactioncode'] == cells[2].text].empty:
                positionIndex = positionIndex + 1
                continue

            new_row = {'Index': positionIndex, 'Datum': cells[0].text, 'Name': '', 'Ebayname': '', 'Produkt': '', 'Mitteilung': '', 'Brutto': brutto, 'Netto': netto, 'Transactioncode': cells[2].text, 'Type': '', 'Bearbeitet': False}
            new_row_df = pd.DataFrame([new_row])
            df_paypal = pd.concat([df_paypal, new_row_df], ignore_index=True)
            positionIndex = positionIndex + 1
        
    df_paypal.to_csv(file_paypal, index=False, sep=';')
    
    return df_paypal

def fill_df_amazon(df_amazon):
    
    shadow_root = expand_shadow_element(driverAmazon.find_element(By.XPATH,'//*[@id="root"]/div/article[3]/section[3]/div/kat-card/div/div[2]/div[1]/kat-pagination'), driverAmazon)
    pages = shadow_root.find_elements(By.CSS_SELECTOR,'li[part^=pagination-page-]')

    shadow_root_arrow = expand_shadow_element(shadow_root.find_element(By.CSS_SELECTOR, 'kat-icon[name=chevron-right]'), driverAmazon)
    nextbutton = shadow_root_arrow.find_element(By.CSS_SELECTOR,'i')
    

    positionIndex = 1
    for page in pages:
        time.sleep(1)
        AmazonRowList = driverAmazon.find_elements(By.XPATH,'//*[@id="root"]/div/article[3]/section[3]/div/kat-card/div/div[1]/kat-table/kat-table-body/kat-table-row')
        AmazonTable = driverAmazon.find_element(By.XPATH,'//*[@id="root"]/div/article[3]/section[3]/div/kat-card/div/div[1]/kat-table/kat-table-body')
        shadow_root = expand_shadow_element(driverAmazon.find_element(By.XPATH,'//*[@id="root"]/div/article[3]/section[3]/div/kat-card/div/div[2]/div[1]/kat-pagination'), driverAmazon)
        shadow_root_arrow = expand_shadow_element(shadow_root.find_element(By.CSS_SELECTOR, 'kat-icon[name=chevron-right]'), driverAmazon)
        nextbutton = shadow_root_arrow.find_element(By.CSS_SELECTOR,'i')
        
        rowcounter = 1
        for transaction in AmazonRowList:
            a = AmazonTable.find_element(By.XPATH,'//*[@id="root"]/div/article[3]/section[3]/div/kat-card/div/div[1]/kat-table/kat-table-body/kat-table-row[' + str(rowcounter) + ']')
            cells = a.find_elements(By.XPATH,'kat-table-cell')


            if (cells[2].text != 'Nicht verfügbarer Saldo aus vorausgehender Abrechnung' and
                cells[2].text != 'Nicht verfügbarer Saldo' and
                cells[2].text != 'Service-Gebühren' and
                cells[2].text != 'Bei Amazon gekaufte Versandetiketten'):

                if cells[2].text != 'Erstattung':
                    transactionnumber = cells[3].text
                else:
                    transactionnumber = cells[3].text + ' Erstattung'
                date = cells[0].text
                produkt = cells[4].text
                netto = float(cells[5].text.replace('.', '').replace(',', '.').replace('€','').replace('$',''))

                if netto > 50 or netto < -50:
                    if not df_amazon[df_amazon['Transactioncode'] == transactionnumber].empty:
                        positionIndex = positionIndex + 1
                        continue
                    new_row = {'Index': positionIndex, 'Datum': date, 'Name': '', 'Produkt': produkt, 'Mitteilung': '', 'Brutto': '', 'Netto': netto, 'Transactioncode': transactionnumber, 'Type': cells[2].text, 'Bearbeitet': False}
                    new_row_df = pd.DataFrame([new_row])

                    df_amazon = pd.concat([df_amazon, new_row_df], ignore_index=True)
                    positionIndex = positionIndex + 1
            
            rowcounter = rowcounter + 1
        df_amazon.to_csv(file_amazon, index=False, sep=';')
        nextbutton.click()
    return df_amazon

def fill_df_ebay(df_ebay):

    #driverEbay.find_element(By.CSS_SELECTOR, '#transactions > div.title-section > div.exclusion-fee-type > span.checkbox > input').click()
    driverEbay.find_element(By.XPATH, '//*[@id="transactions"]/div[1]/div[2]/span[1]/input').click()
    time.sleep(2)
    pages = driverEbay.find_element(By.XPATH, '//*[@id="transactions"]/div[1]/div[1]/span/span/span').text
    vonbis = [int(num) for num in re.findall(r'\d+', pages)]


    positionIndex = 0
    while 1:
        print(vonbis)
        time.sleep(3)
        pages = driverEbay.find_element(By.XPATH, '//*[@id="transactions"]/div[1]/div[1]/span/span/span').text
        vonbis = [int(num) for num in re.findall(r'\d+', pages)]

        ebayPageRows = driverEbay.find_elements(By.XPATH, '//*[@id="transactions"]/div[2]/div')

        for row in ebayPageRows:
            bestllungORrueckzahlung = row.find_element(By.CSS_SELECTOR, 'span[class="BOLD"]').text

            if bestllungORrueckzahlung.startswith('Bestellung') or bestllungORrueckzahlung.startswith('Rückerstattung'):
                date = row.find_element(By.CSS_SELECTOR, 'div.transactions-date > span > span:nth-child(1) > span').text

                transactionnumber = row.find_element(By.CSS_SELECTOR, 'a[class="eui-text-span INLINE_LINK"]').text
                produkt = row.find_element(By.CSS_SELECTOR, 'span[class="eui-text-span"]').text
                print(row.find_element(By.CSS_SELECTOR, 'div.transaction--net > span > div > div > span > span > span.BOLD.each-as-row').text.replace('.', '').replace(',', '.').replace('€','').replace('$','').replace('£','').replace('CA',''))
                preis = float(row.find_element(By.CSS_SELECTOR, 'div.transaction--net > span > div > div > span > span > span.BOLD.each-as-row').text.replace('.', '').replace(',', '.').replace('€','').replace('$','').replace('£','').replace('CA',''))
                name = row.find_element(By.CSS_SELECTOR, 'div.transaction--desc > div.buyer-parent > div > span > span:nth-child(1) > span').text

                if preis > 50 or preis < -50:
                    if df_paypal[df_paypal['Transactioncode'] == transactionnumber].empty:
                        
                        #convert date from 01. Jan. 2021 to 01.01.2021
                        day, month_name, year = date.split('.')
                        month = german_month_to_number(month_name.strip())
                        date = f"{day}.{month}.{year}"
                        
                        new_row = {'Index': positionIndex, 'Datum': date, 'Name': name, 'Ebayname': '', 'Produkt': produkt, 'Mitteilung': '', 'Brutto': preis, 'Netto': preis, 'Transactioncode': transactionnumber, 'Type': bestllungORrueckzahlung, 'Bearbeitet': False}
                        print (new_row)
                        new_row_df = pd.DataFrame([new_row])
                        df_ebay = pd.concat([df_ebay, new_row_df], ignore_index=True)
    
                    positionIndex = positionIndex + 1
        time.sleep(5)
        df_ebay.to_csv(file_ebay, index=False, sep=';')
        if vonbis[1] != vonbis[2]:
            driverEbay.find_element(By.XPATH, '//*[@id="transactions"]/div[3]/div[1]/nav/button[2]').click()
        else:
            break
    
                
    return df_ebay



def buchhaltung_paypal(df_paypal):
    print("BEARBEITUNG PAYPAL\n\n")

    for index, row in df_paypal.iterrows():
        
        if row.empty or row['Bearbeitet'] == False:
            
            print("Die Zeile mit dem index", index, "und Transactioncode", row['Transactioncode'], "wurde noch nicht bearbeitet.")

            p = PaypalPosition(row['Index'], row['Datum'], row['Brutto'], row['Netto'], row['Transactioncode'])
            p.bearbeitung()
            
        else:

            print("Die Zeile mit dem Transactioncode", row['Transactioncode'], "wurde bearbeitet.")
    
    return df_paypal

def buchhaltung_amazon(df_amazon):
    print("BEARBEITUNG EBAY\n\n")

    for index, row in df_amazon.iterrows():
        
        if row.empty or row['Bearbeitet'] == False:
            
            print("Die Zeile mit dem index", index, "und Transactioncode", row['Transactioncode'], "wurde noch nicht bearbeitet.")
            
            a = AmazonPosition(row['Index'], row['Datum'], row['Produkt'], row['Netto'], row['Type'], row['Transactioncode'])
            a.bearbeitung()
            
        else:
            print("Die Zeile mit dem Transactioncode", row['Transactioncode'], "wurde bearbeitet.")
    
    return df_amazon

def buchhaltung_ebay(df_Ebay):
    print("BEARBEITUNG EBAY\n\n")

    for index, row in df_Ebay.iterrows():
        
        if row.empty or row['Bearbeitet'] == False:
            
            print("Die Zeile mit dem index", index, "und Transactioncode", row['Transactioncode'], "wurde noch nicht bearbeitet.")
            
            e = EbayPosition(row['Index'], row['Datum'], row['Name'], row['Produkt'], row['Brutto'], row['Netto'], row['Type'], row['Transactioncode'])
            e.bearbeitung()
            
        else:
            print("Die Zeile mit dem Transactioncode", row['Transactioncode'], "wurde bearbeitet.")
    
    return df_Ebay







# - - M A I N - - #

file_paypal = "Paypal.csv"
file_amazon = "Amazon.csv"
file_ebay = "Ebay.csv"

df_paypal = create_df_csv(file_paypal)
df_amazon = create_df_csv(file_amazon)
df_ebay = create_df_csv(file_ebay)

start_everything()

x=0
if int(input("FILL TABLES? CSV COMPLETE? 1 YES 0 NO:\n")):
    x=1


if x:
    print("FILLING TABLES\n")
    df_amazon = fill_df_amazon(df_amazon)
    #df_paypal = fill_df_paypal(df_paypal)
    #df_ebay = fill_df_ebay(df_ebay)


#df_paypal = buchhaltung_paypal(df_paypal)
df_amazon = buchhaltung_amazon(df_amazon)
#df_ebay = buchhaltung_ebay(df_ebay)

df_paypal.to_csv(file_paypal, index=False, sep=';')
df_amazon.to_csv(file_amazon, index=False, sep=';')
df_ebay.to_csv(file_ebay, index=False, sep=';')

# Close the browser
input("Close?\n")
driverPaypal.quit()
driverLexware.quit()
driverAmazon.quit()
driverEbay.quit()
print("FINISHED")



