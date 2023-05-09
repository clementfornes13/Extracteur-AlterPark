import PyPDF2
from pdf2image import convert_from_path
import cv2
import pytesseract
import openpyxl
from PySimpleGUI import theme, Text, Button, Input, FileBrowse, FolderBrowse, Exit, Image, WIN_CLOSED, popup, Window
from os import path, startfile, remove
from re import findall

class AlterParkExtractor:
    def __init__(s):
        

        theme('Material1')
        s.IMG_PATH = path.join(path.dirname(__file__), 'images')
        s.icon = path.join(s.IMG_PATH, 'Icone.ico')
        s.logo = path.join(s.IMG_PATH, 'Logo TEA FOS.png')
        
    def run(s):
        
        
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
                    pdf_file = open(emplacement_pdf, 'rb') 
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    sheet.title = "AlterPark"
                    sheet['A1'] = 'Numéro de réservation'
                    sheet['B1'] = 'Immatriculation'
                    sheet['C1'] = 'Date 1'
                    sheet['D1'] = 'Date 2'
                    sheet['E1'] = 'Montant total'
                    sheet['F1'] = 'Apporteur'
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        image = convert_from_path('C:/Users/clement.fornes/Desktop/ScriptsPython/Gitrepository/Extracteur-AlterPark/document.pdf', first_page=page_num+1, last_page=page_num+1)[0]
                        image.save(f'page_{page_num+1}.jpg') #Converti en jpg et save
                        img = cv2.imread(f'page_{page_num+1}.jpg', cv2.IMREAD_GRAYSCALE) #Lis l'image
                        remove(f'page_{page_num+1}.jpg')
                        test_reverse = [(1310, 150, 150, 100)] #Rectangle qui permet de savoir si c'est inversé ou non
                        coordonnees = [
                            (225, 70, 640, 80),  # Deplacement x , Deplacement y, Longueur, Largeur (Point en haut a gauche)
                            (1120, 290, 400, 60),
                            (190, 1180, 200, 50),
                            (950, 1180, 200, 50),
                            (1280, 1800, 240, 50),
                            (1200, 1930, 300, 50)
                        ]
                        for x, y, w, h in test_reverse:
                            test = img[y:y+h, x:x+w]
                            pixel_img = (test == 50).sum()
                            print(f"Il y a {pixel_img} pixels d'image dans le rectangle")
                            p=0
                            o=0
                            if pixel_img >= 5:
                                for x, y, w, h in coordonnees:
                                    coord_rect = img[y:y+h, x:x+w]
                                    detection = cv2.rectangle(img, (x, y), (x+w, y+h), (0, 0, 255), 1)
                                    config = '--psm 3 --oem 3'
                                    text = pytesseract.image_to_string(coord_rect, config=config)
                                    print(f'Page {page_num+1}, Coordonnees : ({x}, {y}, {w}, {h}): {text}')
                                    p=p+1
                                    column_num = p  # assuming p=3 in this case
                                    column_letter = openpyxl.utils.get_column_letter(column_num)
                                    sheet[column_letter + str(page_num+1)] = text
                                    #je dois save dans la colonne +1 a chaque fois. donc faire un c+= 
                                    #same pour le else quand l'image est tournée
                                    #chaque valeur en dessous max_row
                                    #voir doc pour voir comment marquer sheet[column=row= etc..]
                            else:
                                img_rota = cv2.rotate(img, cv2.ROTATE_180)
                                for x, y, w, h in coordonnees:
                                    coord_rect = img_rota[y:y+h, x:x+w]
                                    detection = cv2.rectangle(img_rota, (x, y), (x+w, y+h), (0, 0, 255), 1)
                                    config = '--psm 3 --oem 3'
                                    text = pytesseract.image_to_string(coord_rect, config=config)
                                    print(f'Page {page_num+1}, Coordonnees : ({x}, {y}, {w}, {h}): {text}')
                                    o=o+1
                                    column_num = o  # assuming p=3 in this case
                                    column_letter = openpyxl.utils.get_column_letter(column_num)
                                    sheet[column_letter + str(page_num+1)] = text

                        #cv2.imwrite(f'result_page_{page_num+1}.jpg', detection)
                        #cv2.waitKey(0)
                        #cv2.destroyAllWindows()
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

if __name__ == '__main__':
    extracteur = AlterParkExtractor()
    extracteur.run()