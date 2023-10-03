# Script creazione excel  fatturazione nextcharge
# # e relativo invio agli operatori tramite mail

# 14/09/2022, Jack
# 1.calcolo intervallo di date in cui devo inviare i report
# 2.tramite selenium scarico excel
# 3.porto il file nel path corretta
# 4.recupero tutte le righe del .xlsx
# 5.ordino le righe
# 6.per ogni riga recupero tramite auth key id e operatore
# 7.creo dizionario nome operatore / email
# 8.creo excel per ogni operatore
# 9.somma totale per operatore e la inserisco in fondo
# 10.trasformo excel in html e poi pdf
# 11.invio mail con il file
# 12.creo excel gigante con ogni riga da inviare alla Cristina
# per avere un resoconto del mese
# 13.invio mail ad admin
# 14.cancello tutti i file, tranne il .py

# INFO UTILI:
# IL CHROMEDRIVER presente all'interno della cartella nexthcharge
# è la versione per il mac.

# FARE ATTENZIONE A IMPOSTARE LA VARIABILE OGGI,
#  tramite questa scarica i report

# da settare anche il percorso di dove scarica il file,
# contenuto nella variabile: list_of_files

# API TEMPORANEE! da cambiare quando tutto sarà migrato!


import datetime
import glob
import locale
import os
import smtplib
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import dateutil.relativedelta
import numpy
import numpy as np
import pandas
import pdfkit
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# calcolo della data di inzio e della di fine in cui devo scaricare i report!
# OGGI se viene lanciato lo script il primo giono del mese,
# oppure settare a mano il primo giorno del mese

oggi = datetime.datetime(2023, 10, 1)
# oggi = datetime.datetime.now()
oggistr = oggi.strftime("%Y-%m-%d")
# primo giorno del mese precedente
primaData = oggi + dateutil.relativedelta.relativedelta(months=-1)
primaData = primaData.strftime("%Y-%m-%d")
# giorno precedente al giorno indicato,
# che coincide con l'ultimo giorno del mese precedente
ultimaData = oggi
ultimaData = ultimaData.strftime("%Y-%m-%d")
print(
    "data di oggi:" + oggistr + "  data start:" + primaData + "  data stop:" + ultimaData
)
# SELENIUM SCARICO DELL'EXCEL!
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://chargepoint.management/lockscreen")
username = driver.find_element(By.ID, "userId")

password = driver.find_element(By.ID, "password")

username.send_keys("emotion-g.f.")

password.send_keys("lUchilANDoWN")

button = driver.find_element(By.CLASS_NAME, "loginButton")
button.click()

element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "modalUpdates"))
)
driver.get("https://chargepoint.management/nextcharge-revenues")
dataStart = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "formPickerRevenuesExcelStart"))
)
datepicker = driver.find_elements(By.CLASS_NAME, "input-group-addon")
datepicker[0].click()
time.sleep(10)
dataStop = driver.find_element(By.ID, "formPickerRevenuesExcelStop")
dataStart.clear()
dataStart.send_keys(primaData)
dataStop.clear()
dataStop.send_keys(ultimaData)
time.sleep(10)
b = driver.find_element(By.ID, "buttonExportRevenuesExcel")
b.click()
time.sleep(10)
driver.implicitly_wait(100)

# SPOSTO IL FILE NEL PATH CORRETTO
# path dove scarica il file!!!
list_of_files = glob.glob("/Users/giacomofiorucci/Downloads/*")
latest_file = max(list_of_files, key=os.path.getctime)
os.rename(
    latest_file,
    "/Users/giacomofiorucci/Sviluppo/nextcharge/file_di_partenza.xlsx",
)

locale.setlocale(locale.LC_ALL, "it_IT.UTF-8")
operators = []
total = pandas.DataFrame()
excelFiles = glob.glob("*.xlsx")
excelFile = excelFiles[0]
excel_data = pandas.read_excel(excelFile, sheet_name="Ricavo")
html = """
    <!DOCTYPE html>
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">

    <head>
        <title></title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch><o:AllowPNG/></o:OfficeDocumentSettings></xml><![endif]-->
        <style>
            * {
                box-sizing: border-box;
            }

            body {
                margin: 0;
                padding: 0;
            }

            a[x-apple-data-detectors] {
                color: inherit !important;
                text-decoration: inherit !important;
            }

            #MessageViewBody a {
                color: inherit;
                text-decoration: none;
            }

            p {
                line-height: inherit
            }

            .desktop_hide,
            .desktop_hide table {
                mso-hide: all;
                display: none;
                max-height: 0px;
                overflow: hidden;
            }

            @media (max-width:720px) {
                .desktop_hide table.icons-inner {
                    display: inline-block !important;
                }

                .icons-inner {
                    text-align: center;
                }

                .icons-inner td {
                    margin: 0 auto;
                }

                .image_block img.big,
                .row-content {
                    width: 100% !important;
                }

                .mobile_hide {
                    display: none;
                }

                .stack .column {
                    width: 100%;
                    display: block;
                }

                .mobile_hide {
                    min-height: 0;
                    max-height: 0;
                    max-width: 0;
                    overflow: hidden;
                    font-size: 0px;
                }

                .desktop_hide,
                .desktop_hide table {
                    display: table !important;
                    max-height: none !important;
                }
            }
        </style>
    </head>

    <body style="background-color: #ffffff; margin: 0; padding: 0; -webkit-text-size-adjust: none; text-size-adjust: none;">
        <table class="nl-container" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #ffffff;">
            <tbody>
                <tr>
                    <td>
                        <table class="row row-1" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                            <tbody>
                                <tr>
                                    <td>
                                        <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 700px;" width="700">
                                            <tbody>
                                                <tr>
                                                    <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-left: 10px; padding-right: 10px; vertical-align: top; padding-top: 10px; padding-bottom: 10px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                        <table class="image_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                            <tr>
                                                                <td class="pad" style="width:100%;padding-right:0px;padding-left:0px;">
                                                                    <div class="alignment" align="center" style="line-height:10px"><img class="big" src="https://emotion-team.com/wp-content/uploads/2021/01/PNG-EMO-LOGO.png" style="display: block; height: auto; border: 0; width: 374px; max-width: 100%;" width="374"></div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        <table class="divider_block block-2" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                            <tr>
                                                                <td class="pad">
                                                                    <div class="alignment" align="center">
                                                                        <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                            <tr>
                                                                                <td class="divider_inner" style="font-size: 1px; line-height: 1px; border-top: 0px solid #BBBBBB;"><span>&#8202;</span></td>
                                                                            </tr>
                                                                        </table>
                                                                    </div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <table class="row row-2" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                            <tbody>
                                <tr>
                                    <td>
                                        <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #efefef; color: #000000; width: 700px;" width="700">
                                            <tbody>
                                                <tr>
                                                    <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                        <table class="paragraph_block block-2" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                            <tr>
                                                                <td class="pad" style="padding-left:30px;padding-top:25px;">
                                                                    <div style="color:#101112;font-size:16px;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-weight:400;line-height:120%;text-align:left;direction:ltr;letter-spacing:0px;mso-line-height-alt:19.2px;">
                                                                        <p style="margin: 0; margin-bottom: 16px;">Gentile Cliente,</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">in allegato trova il calcolo relativo alle ricariche erogate da <strong>NextCharge</strong> del mese appena trascorso con indicazioni su importo da fatturare.</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;"><strong>Inviare fattura a: amministrazione@emotion-team.com</strong><br role="presentation"><strong>Con oggetto della mail 'FATTURA RICARICHE MENSILI NEXTCHARGE'</strong></p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">Per ulteriori informazioni non esitate a contattarci.</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">Cordiali saluti,<br>Emotion-team</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">&nbsp;</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;"><strong>I nostri riferimenti</strong></p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">Sede legale: Via Gallipoli, 51 - 73013 - Galatina (Lecce)</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">Sede operativa: Via G. Verdi, 24 - 06073 - Corciano (Perugia)</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">Telefono: 075 9280204</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">&nbsp;</p>
                                                                        <p style="margin: 0; margin-bottom: 16px;"><strong>Assistenza</strong></p>
                                                                        <p style="margin: 0; margin-bottom: 16px;">Telefono: 075 9280205</p>
                                                                        <p style="margin: 0;">Email: assistenza@emotion-team.com</p>
                                                                    </div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        <table class="divider_block block-3" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                            <tr>
                                                                <td class="pad">
                                                                    <div class="alignment" align="center">
                                                                        <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                            <tr>
                                                                                <td class="divider_inner" style="font-size: 1px; line-height: 1px; border-top: 0px solid #BBBBBB;"><span>&#8202;</span></td>
                                                                            </tr>
                                                                        </table>
                                                                    </div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        <table class="button_block block-5" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                            <tr>
                                                                <td class="pad" style="text-align:center;padding-top:25px;padding-bottom:50px;">
                                                                    <div class="alignment" align="center">
                                                                        <!--[if mso]><v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" href="https://emotion-team.com/" style="height:42px;width:98px;v-text-anchor:middle;" arcsize="10%" stroke="false" fillcolor="#0068a5"><w:anchorlock/><v:textbox inset="0px,0px,0px,0px"><center style="color:#ffffff; font-family:Arial, sans-serif; font-size:16px"><![endif]--><a href="https://emotion-team.com/" target="_blank" style="text-decoration:none;display:inline-block;color:#ffffff;background-color:#0068a5;border-radius:4px;width:auto;border-top:0px solid transparent;font-weight:400;border-right:0px solid transparent;border-bottom:0px solid transparent;border-left:0px solid transparent;padding-top:5px;padding-bottom:5px;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;text-align:center;mso-border-alt:none;word-break:keep-all;"><span style="padding-left:5px;padding-right:5px;font-size:16px;display:inline-block;letter-spacing:normal;"><span dir="ltr" style="word-break: break-word; line-height: 32px;">VAI AL SITO</span></span></a>
                                                                        <!--[if mso]></center></v:textbox></v:roundrect><![endif]-->
                                                                    </div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <table class="row row-3" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                            <tbody>
                                <tr>
                                    <td>
                                        <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 700px;" width="700">
                                            <tbody>
                                                <tr>
                                                    <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                        <table class="icons_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                            <tr>
                                                                <td class="pad" style="vertical-align: middle; color: #9d9d9d; font-family: inherit; font-size: 15px; padding-bottom: 5px; padding-top: 5px; text-align: center;">
                                                                    <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                        <tr>
                                                                            <td class="alignment" style="vertical-align: middle; text-align: center;">
                                                                                <!--[if vml]><table align="left" cellpadding="0" cellspacing="0" role="presentation" style="display:inline-block;padding-left:0px;padding-right:0px;mso-table-lspace: 0pt;mso-table-rspace: 0pt;"><![endif]-->
                                                                                <!--[if !vml]><!-->
                                                                                <table class="icons-inner" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; display: inline-block; margin-right: -4px; padding-left: 0px; padding-right: 0px;" cellpadding="0" cellspacing="0" role="presentation">
                                                                                    <!--<![endif]-->
                                                                                    <tr>
                                                                                        <td style="vertical-align: middle; text-align: center; padding-top: 5px; padding-bottom: 5px; padding-left: 5px; padding-right: 6px;"><a href="https://www.designedwithbee.com/?utm_source=editor&utm_medium=bee_pro&utm_campaign=free_footer_link" target="_blank" style="text-decoration: none;"><img class="icon" alt="Designed with BEE" src="https://d15k2d11r6t6rl.cloudfront.net/public/users/Integrators/BeeProAgency/53601_510656/Signature/bee.png" height="32" width="34" align="center" style="display: block; height: auto; margin: 0 auto; border: 0;"></a></td>
                                                                                        <td style="font-family: Arial, Helvetica Neue, Helvetica, sans-serif; font-size: 15px; color: #9d9d9d; vertical-align: middle; letter-spacing: undefined; text-align: center;"><a href="https://www.designedwithbee.com/?utm_source=editor&utm_medium=bee_pro&utm_campaign=free_footer_link" target="_blank" style="color: #9d9d9d; text-decoration: none;">Designed with BEE</a></td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table><!-- End -->
    </body>

    </html>
    """

# dizionario per ricavare email
dizionario = {
    "Magrelli Ospitalità S.r.l.": "amministrazione@magrelli.com",
    "Camer E-Mobility": "emobility@camerpetroleum.it",
    "Emotion S.r.l.": "sviluppatori@emotion-team.com",
    "Euroenergia": "granata_c@libero.it",
    "ASMTERNI SPA": "fabio.paoletti@asmterni.it",
    "Emotion Fast": "assistenza@emotion-team.com",
    "CONSAUTO SOC. COOP.": "amministrazione@consauto.it",
    "Blu oil srl": "alessio.sommacal@azzalinienergie.it",
    "Marche Smart Mobility srl": "info@masmomobility.it",
    "EG ITALIA S.p.A.": "italy.fuelpricingroma@eg.group",
}
# dizionario di prova
dizionario_prova = {
    "Camer E-Mobility": "fgiacomo92@gmail.com",
    "Emotion S.r.l.": "cpncrs@gmail.com",
    "Euroenergia": "fgiacomo92@hotmail.it",
    "ASMTERNI SPA": "giacomo.fiorucci@emotion-team.com",
    "Emotion Fast": "assistenza@emotion-team.com",
}


columnsTitles = [
    "Nome stazione:",
    "Id transazione",
    "User Id",
    "Ricavo",
    "Costo",
    "Metodo",
    "kWh",
    "Data inizio",
    "Data fine",
    "Country",
    "Connettore",
    "Charge Box Identity",
    "Id Pay Transaction",
]


# filtro per una colonna (ad esempio nome stazione) e lo salvo su una variabile
nomeStazione = excel_data.filter(items=["Nome stazione:"])
nomeStazione = nomeStazione.drop_duplicates()
nomeStazione.columns = [""] * len(nomeStazione.columns)
print(nomeStazione)


# riordino le colonne in base columnsTitles
excel_data = excel_data.reindex(columns=columnsTitles)
# riordino le righe
excel_data = excel_data.sort_values(
    by=["Charge Box Identity", "Data inizio", "Nome stazione:"]
)

# somme parziali al momento per colonnina! poi per operatore
# excel_data_partial = (
#    excel_data["Costo"].groupby(excel_data["Nome stazione:"]).sum()
# )

# recupero gli operatori per ogni colonnina!

# filtro per una colonna (Charge Box Identity) e lo salvo su una variabile
chargeboxId = excel_data.filter(items=["Charge Box Identity"])
# chargeboxId = chargeboxId.drop_duplicates()
chargeboxId.columns = [""] * len(chargeboxId.columns)
print(chargeboxId)
# commissioni spotlink da sottrarre
commissioniSpotlink = 0.02
# calcolo finale costo
excel_data["Totale Ricarica"] = (
    excel_data["Ricavo"] / excel_data["kWh"] - commissioniSpotlink
) * excel_data["kWh"]


chargeboxId = chargeboxId.to_numpy()
chargeboxId = chargeboxId.ravel()

# funzione che mi permette di cancellare eventuali colonne da non mostrare
excel_data = excel_data.drop(columns=["Metodo"])
excel_data = excel_data.drop(columns=["kWh"])
excel_data = excel_data.drop(columns=["Country"])
excel_data = excel_data.drop(columns=["Ricavo"])
excel_data = excel_data.drop(columns=["Costo"])
excel_data = excel_data.drop(columns=["Connettore"])
excel_data = excel_data.drop(columns=["Id Pay Transaction"])
excel_data = excel_data.drop(columns=["Id transazione"])


for n in chargeboxId:
    # CHIAMATA API, IN FUTURO ANDRà MODIFICATA!!!!!!!
    """
    response = requests.get(
        "https://emotion-projects.eu/api/towers/",
        params={"nextcharge_auth_key": n},
    )
    print(n)
    json = response.json()
    print(json)
    results = json["results"]
    operator = results["operator"]
    """
    dizionario_operatori = {
        "JZ5I4XRWH10F9PRZ24WD": "Magrelli Ospitalità S.r.l.",
        "EEPJFFR75BK5CA6YY3WR": "Magrelli Ospitalità S.r.l.",
        "HTOM9S8494IGT9IQDKST": "Camer E-Mobility",
        "JEUW0GCL613UUZMYGLQW": "Camer E-Mobility",
        "71Z1162H9KAURF2X8LJK": "Camer E-Mobility",
        "HWPQMQ2G8ILFIKVXBSCT": "Camer E-Mobility",
        "OBA2EHJKMI84Z1XNWQZU": "Camer E-Mobility",
        "JPXOBFU9YO0S7OKSKKFZ": "Emotion S.r.l.",
        "WD3YAS4ZMZV8N2ST3986": "Camer E-Mobility",
        "3CFFK6OYRSN1WBQD9E3R": "Emotion S.r.l.",
        "QJGXGLMHQPXDQ5XCMLUS": "Emotion S.r.l.",
        "QRUSATX4F1UTCWID6AUZ": "CONSAUTO SOC. COOP.",
        "GBWSSW8C6EPWC3ADTWJY": "Camer E-Mobility",
        "SE2UC2TN1EHZCX8TQGXK": "Camer E-Mobility",
        "UCUU5ABGBA0X8I7OMSZW": "Camer E-Mobility",
        "FMG0JXWR7ZIWXZ4WNRES": "Camer E-Mobility",
        "JXH3S4UMSX9YX4I2MO6C": "Camer E-Mobility",
        "EMT6NTQ12I29V4YGGFQ7": "Camer E-Mobility",
        "XTWQL4WPZ5AOBA3HIQKZ": "Camer E-Mobility",
        "VAVZNU3MEBV3G63MP2A6": "Camer E-Mobility",
        "4PBX7VKDUAMD2QEG9S4I": "Camer E-Mobility",
        "HL3N1I5551FAY5JWVPX7": "Camer E-Mobility",
        "FCK9792PZ2ZJEB9S4ZLM": "CONSAUTO SOC. COOP.",
        "L6FP96YBRSKKY3SH538S": "CONSAUTO SOC. COOP.",
        "FGQE44SA1OOVTYSSV1DJ": "Emotion Fast",
        "JOGYCCEQ80DWMEZOAX2Z": "ASMTERNI SPA",
        "LMFXG0JXYQUHA9FV9879": "Emotion S.r.l.",
        "4YG3FU1V4NPZN0AWRQD6": "Blu oil srl",
        "WMHLXF6EBNLXEIWNMVM5": "Marche Smart Mobility srl",
        "O8K3MTVKZXOH4G3TRS27": "EG ITALIA S.p.A.",
        "TMO3HXBCJGAXG09EO726": "EG ITALIA S.p.A.",
        "3F10X2MXW0RWCQBLU14Q": "EG ITALIA S.p.A.",
    }
    operator = dizionario_operatori[n]
    operators.append(operator)
print(operators)
operatorsList = np.array(operators)
print(operatorsList)
df = excel_data
df["operatori"] = operatorsList.tolist()
# df = pandas.concat([excel_data, # ], axis=1)
# lista senza duplicati
operators = list(dict.fromkeys(operators))


for n in operators:
    ricariche_per_singolo_operatore = df[df["operatori"] == n]
    ricariche_per_singolo_operatore.at[
        " Totale da Fatturare", "Totale Ricarica"
    ] = ricariche_per_singolo_operatore["Totale Ricarica"].sum()
    nomeripulito = n.replace(" ", "").replace(".", "").replace("-", "")
    nomeExcel = pandas.ExcelWriter(nomeripulito + ".xlsx", engine="xlsxwriter")
    ricariche_per_singolo_operatore.to_excel(
        nomeExcel, sheet_name="Riepilogo Ricariche", index=False
    )
    workbook = nomeExcel.book
    worksheet = nomeExcel.sheets["Riepilogo Ricariche"]
    for col_idx, col in enumerate(df.columns):
        max_len = (
            max(df[col].astype(str).str.len().max(), len(col)) + 1
        )  # Aggiungi spazio extra
        worksheet.set_column(col_idx, col_idx, max_len)

    # nomeExcel.save()

    total = pandas.concat([total, ricariche_per_singolo_operatore], ignore_index=True)

    ricariche_per_singolo_operatore.to_html(
        nomeripulito + ".html",
        header=True,
        index=True,
        na_rep="",
        float_format="{:20,.2f}".format,
    )
    nomehtml = nomeripulito + ".html"
    nomePdf = nomeripulito + ".pdf"
    print("nome pdf!!!!!  " + nomePdf)
    pdfkit.from_file(nomehtml, nomePdf)
    fileDaInviare = nomePdf
    print(fileDaInviare)

    # invio mail
    msg = MIMEMultipart("mixed")
    msg["Subject"] = "Fatture NextCharge"
    msg["From"] = "notifiche@spot-link.it"
    msg["To"] = dizionario[n]
    msg["CC"] = "assistenza@emotion-team.com"

    text = MIMEText(html, "html")
    msg.attach(text)
    attachmentPath = fileDaInviare
    try:
        with open(attachmentPath, "rb") as attachment:
            p = MIMEApplication(attachment.read(), _subtype="xlsx")
            p.add_header(
                "Content-Disposition",
                "attachment; filename= %s" % attachmentPath.split("\\")[-1],
            )
            msg.attach(p)
    except Exception as e:
        print(str(e))

    try:
        smtpObj = smtplib.SMTP_SSL("smtps.aruba.it", 465)
        smtpObj.set_debuglevel(1)
        smtpObj.login("notifiche@spot-link.it", "Sl#notify22")
        smtpObj.sendmail(
            "notifiche@spot-link.it",
            [dizionario[n], "assistenza@emotion-team.com"],
            msg.as_string(),
        )
        print(msg.as_string())
    except smtplib.SMTPException as e:
        print("NOOOOO")
        print(e)
total.to_excel("Riepilogo_Ricariche_Nextcharge.xlsx", sheet_name="Riepilogo Ricariche")
# Invio alla cristina del riepilogo
msg = MIMEMultipart("mixed")
msg["Subject"] = "Riassunto Fatture NextCharge Per Amministrazione"
msg["From"] = "notifiche@spot-link.it"
msg["To"] = "amministrazione@emotion-team.com"
msg["CC"] = "assistenza@emotion-team.com"

text = MIMEText(html, "html")
msg.attach(text)

attachmentPath = "Riepilogo_Ricariche_Nextcharge.xlsx"
try:
    with open(attachmentPath, "rb") as attachment:
        p = MIMEApplication(attachment.read(), _subtype="xlsx")
        p.add_header(
            "Content-Disposition",
            "attachment; filename= %s" % attachmentPath.split("\\")[-1],
        )
        msg.attach(p)
except Exception as e:
    print(str(e))
try:
    smtpObj = smtplib.SMTP_SSL("smtps.aruba.it", 465)
    smtpObj.set_debuglevel(1)
    smtpObj.login("notifiche@spot-link.it", "Sl#notify22")
    smtpObj.sendmail(
        "notifiche@spot-link.it",
        ["amministrazione@emotion-team.com", "assistenza@emotion-team.com"],
        msg.as_string(),
    )
    print(msg.as_string())
except smtplib.SMTPException as e:
    print("NOOOOO")
    print(e)
# somma totale
df.at["Totale", "Totale Ricarica"] = excel_data["Totale Ricarica"].sum()

# df.to_excel("pandas_to_excel.xlsx", sheet_name="new_sheet_name")

# elimino tutti gli excel
excelFiles = glob.glob("*.xlsx")
for f in excelFiles:
    if os.path.exists(f):
        os.remove(f)
    else:
        print("The file does not exist")
# elimino tutti gli html
excelFiles = glob.glob("*.html")
for f in excelFiles:
    if os.path.exists(f):
        os.remove(f)
    else:
        print("The file does not exist")
# elimino tutti i pdf
excelFiles = glob.glob("*.pdf")
for f in excelFiles:
    if os.path.exists(f):
        os.remove(f)
    else:
        print("The file does not exist")
