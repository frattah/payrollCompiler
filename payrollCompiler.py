import numpy as np
import logging
import pytesseract
from pdf2image import convert_from_path
import sys
import cv2
import re
from PIL import Image
from openpyxl import load_workbook
from pathlib import Path

######################################################################################################################
#                                                                                                                    #
#                                               PARAMETERS                                                           #
#                                                                                                                    #
######################################################################################################################

WRITING_PARAMETERS = ["attendances", "vacancies","tickets","0131","holidays","0421","0457"]
CODES = ["0131","0200","0202","0203","0205","0206","0207","0210","0293","0299","0352","0353","0366","0421","0457"]
HOLIDAYS = ["0200","0202","0203","0205","0206","0207","0210","0352","0353","0366"]
TICKETS = ["0293","0299"]
MONTHS = {
    "Gennaio": 1,
    "Febbraio": 2,
    "Marzo": 3,
    "Aprile": 4,
    "Maggio": 5,
    "Giugno": 6,
    "Luglio": 7,
    "Agosto": 8,
    "Settembre": 9,
    "Ottobre": 10,
    "Novembre": 11,
    "Dicembre": 12
}
X_STARTING_CELL = 5
Y_STARTING_CELL = 8
FIRST_YEAR = 2007
excel_path = sys.argv[1]

######################################################################################################################
#                                                                                                                    #
#                                                     CODE                                                           #
#                                                                                                                    #
######################################################################################################################

class Payroll:
    def __init__(self, year, month):
        self.year = year
        self.month = month
        self.pay_elements = {}

    def set_pay_element(self, name, value):
        if name in self.pay_elements:
            self.pay_elements[name] += value
        else:
            self.pay_elements[name] = value

    def get_pay_element(self, name):
        if name not in self.pay_elements:
            return 0
        return self.pay_elements[name]

    def compute_derivates(self):
        self.pay_elements["holidays"] = sum(self.pay_elements[key] for key in self.pay_elements if key in HOLIDAYS)

        # Choose the right ticket value
        if self.year > 2023 or (self.year == 2023 and self.month >= 10):
            ticket_value = 10.50
        elif self.year < 2012 or (self.year == 2012 and self.month < 10):
            ticket_value = 0
        else:
            ticket_value = 7.30

        # Multiply ticket value to ticket parameter
        if "0293" in self.pay_elements:
            self.pay_elements["tickets"] = self.pay_elements["0293"] * ticket_value
        elif "0299" in self.pay_elements:
            self.pay_elements["tickets"] = self.pay_elements["0299"] * ticket_value

    def write_on_spreadsheet(self, sheet):
        for i,pay_element in enumerate(WRITING_PARAMETERS):
            sheet.cell(row=X_STARTING_CELL + (self.year - FIRST_YEAR) * 14 + i, 
                            column=Y_STARTING_CELL + MONTHS[self.month], 
                            value= self.pay_elements[pay_element] if pay_element in self.pay_elements else 0.00 )

# ------------------------------------------------------------------------------------------------------------------

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def process_pdf(pdf_file):

    pages = convert_from_path(pdf_file, dpi=300, poppler_path=r"C:\Users\framo\Downloads\Release-25.07.0-0\poppler-25.07.0\Library\bin")

    # Read pdf content after converting it to image
    text = ""
    for page_num, page in enumerate(pages, start=1):
        img = page.convert("RGB")
        custom_config = r'--psm 4 -l ita'
        page_text = pytesseract.image_to_string(img, config=custom_config)
        text += page_text
    #print(page_text)

    # Find month and year of the payroll
    match = re.search(r"Stipendio(?:.*?\s)?([A-Za-z]+)\s+(\d{4})", text, re.IGNORECASE | re.DOTALL) # Case insensitive and capture also '\n'


    if match:
        month = match.group(1)
        year = int(match.group(2))

    payroll = Payroll(year, month)

    # Find attendance and vacancies
    pattern = r"Ferie anno.*?(\d{1,2},\d{2})(?:\s+\d{1,2},\d{2})\s+(\d{1,2},\d{2})"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if match:
        attendances = float(match.group(1).replace(',','.'))
        vacancies = float(match.group(2).replace(',','.'))
        payroll.set_pay_element("attendances", attendances)
        payroll.set_pay_element("vacancies", vacancies)

    # Find all pay elements in the payroll
    for code in CODES:
        if code in TICKETS:
            # Pattern: codice all'inizio della riga, poi qualsiasi cosa, poi cattura il primo numero con eventuale virgola
            pattern = rf"^{code}\b.*?\s(\d+(?:,\d+)?)"
        else:
            # Per gli altri codici, continua a prendere l'ultimo numero della riga
            pattern = rf"^{code}\b.*?(\d+(?:[.,]\d+)?)$"

        for match in re.finditer(pattern, text, re.MULTILINE):
            value = float(match.group(1).replace(",", "."))
            payroll.set_pay_element(code, value)

    return payroll

def process_payrolls(sheet):
    all_payrolls = []

    for pdf_file in Path(".").rglob("*.pdf"):
        try:
            payroll = process_pdf(str(pdf_file))
            all_payrolls.append(payroll)
            print(f"Processed {payroll.month} {payroll.year}")
        except Exception as e:
            print(f"Errore con {pdf_file}: {e}")
    
    for payroll in all_payrolls:
        payroll.compute_derivates()
        payroll.write_on_spreadsheet(sheet)


wb = load_workbook(excel_path)
sheet = wb[wb.sheetnames[0]]

process_payrolls(sheet)

wb.save("buste_modificato.xlsx")



