import os 
import openpyxl
from string import ascii_uppercase

os.system('cls')

dateiName = "test.xlsx"

def neueExcelTabelleErstellen(dateiName):
    wb = openpyxl.Workbook()

    trainingsSheet = wb.active

    titelListe = ['Uebung', 'Beschreibung', 'Saetze', 'Wiederholungen']

    # Überschriften in die oberste Zeile der Exceltabelle eintragen
    for i in range(0, len(titelListe)):
        trainingsSheet[str(ascii_uppercase[i]) + str(1)] = titelListe[i]

    wb.save(filename=dateiName)
    

neueExcelTabelleErstellen(dateiName)