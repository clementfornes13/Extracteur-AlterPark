from PyPDF2 import PdfReader
from pdf2image import convert_from_path
from cv2 import imread, IMREAD_GRAYSCALE, rotate, ROTATE_180
from pytesseract import image_to_string, pytesseract
from openpyxl import Workbook
from PySimpleGUI import theme, Text, Button, Input, FileBrowse, FolderBrowse, Exit, WIN_CLOSED, popup, Window
from os import path, startfile, remove, add_dll_directory, path
import re
import sys

class AlterParkExtractor:
    def __init__(self):
        theme('Material1')
        self.DATE_PATTERN = r"\d{2}/\d{2}/\d{4}"
        self.IMMAT_PATTERN = r"[A-Z]{2}-[0-9]{3}-[A-Z]{2} | [A-Z]{2}[0-9]{3}[A-Z]{2}"
        self.LIST_APPORTEUR = ['parkcloud', 'travelcar', 'ZENPARK', 'parkos', 'travelercar']
        self.APPORTEUR_PATTERN = r"\b(" + "|".join(self.LIST_APPORTEUR) + r")\b"

    def process_pdf(self, emplacement_pdf):
        pdf_file = open(emplacement_pdf, 'rb')
        pdf_reader = PdfReader(pdf_file)
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "AlterPark"
        sheet['A1'] = 'Numéro de réservation'
        sheet['B1'] = 'Immatriculation'
        sheet['C1'] = 'Date 1'
        sheet['D1'] = 'Date 2'
        sheet['E1'] = 'Montant total'
        sheet['F1'] = 'Apporteur'
        config = '--psm 3 --oem 3'
        for page_num in range(len(pdf_reader.pages)):
            image = convert_from_path(emplacement_pdf, first_page=page_num + 1, last_page=page_num + 1)[0]
            image.save(f'page_{page_num + 1}.jpg')
            img = imread(f'page_{page_num + 1}.jpg', IMREAD_GRAYSCALE)
            remove(f'page_{page_num + 1}.jpg')
            test_reverse = [(1280, 45, 320, 160)]
            for x, y, w, h in test_reverse:
                test = img[y:y + h, x:x + w]
                pixel_img = (test == 50).sum()
                if pixel_img >= 5:
                    texte = image_to_string(img, config=config)
                else:
                    img_rota = rotate(img, ROTATE_180)
                    texte = image_to_string(img_rota, config=config)
                resa = texte[8:15] if texte[8:15].startswith('A') else f'A{texte[8:15]}'
                sheet[f'A{str(page_num + 2)}'] = resa.replace(" ", "").replace("O", "0")
                sheet[f'B{str(page_num + 2)}'] = re.findall(self.DATE_PATTERN, texte)[0]
                sheet[f'C{str(page_num + 2)}'] = re.findall(self.DATE_PATTERN, texte)[1]
                if immat := re.findall(self.IMMAT_PATTERN, texte):
                    sheet[f'D{str(page_num + 2)}'] = immat[0]
                lignes = texte.split("\n")
                for l in lignes:
                    if "TOTAL" in l:
                        montant_total = l.split("TOTAL", 1)[1].strip()
                        if "€TTC" in montant_total:
                            sheet[f'E{str(page_num + 2)}'] = montant_total
                        else:
                            sheet[f'E{str(page_num + 2)}'] = f"{montant_total} €TTC"
                sheet[f'F{str(page_num + 2)}'] = re.findall(self.APPORTEUR_PATTERN, texte, flags=re.IGNORECASE)[0]
        return workbook

    def extraire_data(self, emplacement_pdf, destination):
        if not emplacement_pdf:
            popup('Pas de fichier sélectionné, veuillez en sélectionner un !', title='Erreur', icon=self.icon)
            return
        workbook = self.process_pdf(emplacement_pdf)
        file_name = "Extraction AlterPark.xlsx".format(1)
        file_path = path.join(destination, file_name)
        i = 1
        while path.exists(file_path):
            i += 1
            file_name = f"Extraction AlterPark {i}.xlsx"
            file_path = path.join(destination, file_name)
        workbook.save(file_path)
        popup("Fini!", f"Fichier enregistré ici : {file_path}")
        startfile(file_path)

    def run(self):
        layout = [
            [Text("Fichier PDF :"), Input(), FileBrowse('Parcourir', file_types=(("Fichiers PDF", "*.pdf"),))],
            [Text("Destination de l'extraction :"), Input(key='Destination'), FolderBrowse('Parcourir')],
            [Button("Extraction"), Exit('Quitter')]
        ]
        self.window = Window("Extracteur AlterPark", layout)
        while True:
            event, values = self.window.read()
            if event in [WIN_CLOSED, "Quitter"]:
                break
            elif event == "Extraction":
                self.extraire_data(values[0], values['Destination'])
        self.window.close()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = path.abspath(".")

    return path.join(base_path, relative_path)

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        # Running as a bundled executable
        bundle_dir = sys._MEIPASS
    else:
        # Running in the development environment
        bundle_dir = path.dirname(path.abspath(__file__))
    poppler_path = path.join(bundle_dir, 'poppler', 'bin')
    pytesseract.tesseract_cmd = path.join(bundle_dir, 'tesseract', 'tesseract')
    add_dll_directory(poppler_path)
    extracteur = AlterParkExtractor()
    extracteur.run()
