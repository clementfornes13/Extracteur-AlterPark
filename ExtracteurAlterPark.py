#Auto -generated docs, and comment using a tool named : Mintlify Doc Writer

# The PDFExtractor class extracts specific information from a PDF file and saves it to an Excel file.

from PySimpleGUI import theme, Text, Button, Input, FileBrowse, FolderBrowse, Exit, Image, WIN_CLOSED, popup, Window
from os import path, startfile
from re import findall
from PyPDF2 import PdfReader
from openpyxl import Workbook

class AlterParkExtractor:
    def __init__(s):
        
        """
        The function initializes several regular expression patterns and sets the theme to 'Material1'.
        
        :param s: The parameter "s" is a reference to the instance of the class being created. It is
        commonly used as a convention for the "self" parameter in Python classes. The "__init__" method is a
        special method in Python classes that is called when an instance of the class is created. In
        """
        
        s.NUMERO_RESA_PATTERN = r"A\d{6}"
        s.APPORTEUR_PATTERN = r"\b([A-Z]+)\b"
        s.DATES_PATTERN = r"\d{2}/\d{2}/\d{4}"
        s.TOTAL_PATTERN = r"\d+,\d{2}"
        theme('Material1')
        s.IMG_PATH = path.join(path.dirname(__file__), 'images')
        s.icon = path.join(s.IMG_PATH, 'Icone.ico')
        s.logo = path.join(s.IMG_PATH, 'Logo TEA FOS.png')
        
    def run(s):
        
        """
        This is a Python function that extracts information from a PDF file and saves it to an Excel
        file, with a GUI interface for selecting the input and output files.
        
        :param s: The parameter "s" is likely an instance of a class or a namespace object that contains
        various attributes and methods used in the "run" function. These attributes and methods are
        likely used to define the layout of the GUI window, extract information from PDF files, and save
        the extracted information to an Excel
        """
        
        layout = [
            [Text("Fichier PDF :"), Input(), FileBrowse('Parcourir',file_types=(("Fichiers PDF", "*.pdf"),))],
            [Text("Destination de  l'extraction :"), Input(key='Destination'), FolderBrowse('Parcourir')],
            [Button("Extraction"), Exit('Quitter')]]
        
        s.window = Window("Extracteur AlterPark", layout,icon=s.icon)
        while True:
            event, values = s.window.read()
            if event == WIN_CLOSED or event == "Quitter": 
                break
            elif event == "Extraction": 
                emplacement_pdf = values[0] 
                if emplacement_pdf:
                    fichier_pdf = open(emplacement_pdf, 'rb') 
                    lire_pdf = PdfReader(fichier_pdf) 
                    workbook = Workbook()
                    sheet = workbook.active
                    sheet['A1'] = 'Numéro de résa'
                    sheet['B1'] = 'Date 1'
                    sheet['C1'] = 'Date 2'
                    sheet['D1'] = 'Montant total'
                    sheet['E1'] = 'Apporteur'
                    for page_num in range(len(lire_pdf.pages)):
                        page = lire_pdf.pages[page_num]
                        texte = page.extract_text()
                        print(texte)
                        dates = findall(s.DATES_PATTERN, texte)
                        montanttotal  = findall(s.TOTAL_PATTERN, texte)
                        num_resa = findall(s.NUMERO_RESA_PATTERN, texte)
                        apporteur = findall(s.APPORTEUR_PATTERN, texte)
                        sheet['A' + str(sheet.max_row)] = num_resa[0]
                        sheet['B' + str(sheet.max_row)] = dates[0]
                        sheet['C' + str(sheet.max_row)] = dates[1]
                        sheet['D' + str(sheet.max_row)] = montanttotal[0]
                        sheet['E' + str(sheet.max_row)] = apporteur[15]
                    file_name = "Extraction infos.xlsx".format(1)
                    file_path = path.join(values['Destination'],file_name)
                    i = 1
                    while path.exists(file_path):
                        i += 1
                        file_name = "Extraction {}.xlsx".format(i)
                        file_path = path.join(values['Destination'],file_name)
                    workbook.save(file_path)
                    popup("Fini!", f"Fichier enregistré ici : {file_path}")
                    startfile(file_path)
                else:
                 popup('Pas de fichier sélectionné, veuillez en sélectionner un !',title='Erreur',icon=s.icon)
        s.window.close()
# This code block is checking if the current script is being run as the main program (as opposed to
# being imported as a module into another program). If it is being run as the main program, it creates
# an instance of the AlterParkExtractor class and calls its run() method, which starts the PySimpleGUI
# window and runs the AlterPark extraction program.

if __name__ == '__main__':
    extracteur = AlterParkExtractor()
    extracteur.run()