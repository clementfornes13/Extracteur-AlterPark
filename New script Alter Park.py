import tkinter as tk
from tkinter import filedialog, messagebox
from os import path
from threading import Thread
from fitz import open, Matrix
from PIL import Image
from pytesseract import image_to_string
from re import findall, IGNORECASE
from tkinter import ttk
from datetime import datetime
from openpyxl import Workbook
from time import perf_counter

class PDFFinder:
    def __init__(self, root):
        self.root = root
        self.root.title("Extraire Data Alter Park")
        self.root.geometry("300x330")
        self.root.resizable(False, False)
        self.root.help = True
        self.pdf_file_path = tk.StringVar()
        self.destination = tk.StringVar()

        self.reservation_number_pattern = r"(?<![0-9])0*(\d{5})"
        self.date_pattern = r"(\d{1,2}/\d{1,2}/\d{4})"
        self.total_price_pattern = r"(\d+,\d{2})"
        self.license_plate_pattern =  r"\b(?:[A-Z]{2}\s?\d{3}\s?[A-Z]{2}|[A-Z]{2}\d{2}\s?[A-Z]{2}|[A-Z]{2}\s?\d{3}[A-Z]{2})\b"
        self.license_plate_pattern1 = r'\b[A-Z]{2}\s?[-]?\s?\d{3}\s?[-]?\s?[A-Z]{2}\b|\b\d{3}\s?[-]?\s?[A-Z]{2}\s?[-]?\s?\d{3}\b'
        self.license_plate_pattern2 = r'\b[A-Z]{2}\d{2}\s?[A-Z]{3}\b|\b[A-Z]{2}\d{3}\s?[A-Z]{3}\b'
        self.license_plate_pattern3 = r'\b[A-Z]{3}\s?[-]?\s?\d{4}\b|\b\d{4}\s?[-]?\s?[A-Z]{3}\b'

        self.create_widgets()

    def create_widgets(self):
        
        self.label_pdf_file = tk.Label(self.root, text="Fichier PDF:")
        self.label_pdf_file.pack(pady=5)

        self.entry_pdf_file = tk.Entry(self.root, textvariable=self.pdf_file_path)
        self.entry_pdf_file.pack(pady=5)

        self.btn_browse_pdf = tk.Button(self.root, text="Parcourir", command=self.browse_pdf)
        self.btn_browse_pdf.pack(pady=5)

        self.label_destination = tk.Label(self.root, text="Destination:")
        self.label_destination.pack(pady=5)

        self.entry_destination = tk.Entry(self.root, textvariable=self.destination)
        self.entry_destination.pack(pady=5)

        self.btn_browse_destination = tk.Button(self.root, text="Parcourir", command=self.browse_destination)
        self.btn_browse_destination.pack(pady=5)

        self.btn_start_extraction = tk.Button(self.root, text="Lancer", command=self.start_extraction)
        self.btn_start_extraction.pack(pady=10)

        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)

        self.label_progress = tk.Label(self.root, text="0%")
        self.label_progress.pack(pady=2)
        
        self.label_time_remaining = tk.Label(self.root, text="Temps restant : 00:00:00")
        self.label_time_remaining.pack(pady=2)
        
    def browse_pdf(self):
        if file_path:= filedialog.askopenfilename(
            filetypes=[("Fichiers PDF", 
                        "*.pdf")]
        ):
            self.pdf_file_path.set(file_path)

    def browse_destination(self):
        if folder_path:= filedialog.askdirectory():
            self.destination.set(folder_path)

    def validate_pdf_file_path(self, pdf_file_path):
        # Validate the PDF file path
        if not path.isfile(pdf_file_path):
            messagebox.showerror("Erreur", "Chemin du fichier PDF invalide.")
            return False
        elif not pdf_file_path.lower().endswith(".pdf"):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier PDF.")
            return False
        return True

    def start_extraction(self):
        pdf_file_path = self.pdf_file_path.get()
        destination = self.destination.get()

        if not destination:
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier de destination.")
            return

        if self.validate_pdf_file_path(pdf_file_path):
            thread = Thread(target=self.process_pdf, args=(pdf_file_path, destination))
            thread.start()

    def process_pdf(self, pdf_file_path, destination):
        doc = open(pdf_file_path)
        total_pages = len(doc)

        # Definir les valeurs de recherche
        search_values = ["parkcloud", "travelcar", "zenpark", "parkos", "travelercar"]

        # Compteurs pour les valeurs trouvées
        total_pages_found = 0
        reservation_number_count = 0
        date_depot_count = 0
        date_restitution_count = 0
        total_price_count = 0
        plate_number_count = 0
        apporteur_count = 0
        # Liste pour stocker les pages avec des patterns
        pages_with_patterns = []
        # Créez un nouveau classeur Excel
        workbook = Workbook()

        # Ajoutez une feuille de calcul au classeur
        sheet = workbook.active
        sheet.title = "Data Extraite"

        # Ajoutez les en-têtes des colonnes
        column_headers = [
            "Numero de page",
            "Numero de réservation",
            "Date de dépot",
            "Date de restitution",
            "Montant Total",
            "Immatriculation",
            "Apporteur"
        ]
        sheet.append(column_headers)
        self.start_time = perf_counter()
        for page_num in range(total_pages):
            # Mettre à jour la barre de progression et le pourcentage de progression
            progress_percent = (page_num + 1) / total_pages * 100
            self.progress_var.set(progress_percent)
            self.label_progress.config(text=f"{progress_percent:.2f}%")
            # Calculer le temps restant estimé
            time_remaining = (perf_counter() - self.start_time) / (page_num + 1) * (total_pages - page_num - 1)
            self.label_time_remaining.config(text=f"Temps restant : {time_remaining:.0f} secondes")
            self.root.update()

            page = doc.load_page(page_num)
            dpi_factor = 300
            pix = page.get_pixmap(matrix=Matrix(dpi_factor / 72, dpi_factor / 72))  # Augmenter la résolution de l'image
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Faire l'OCR de l'image en utilisant Tesseract
            ocr_text = image_to_string(image)
            
            """""
            ocr_data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)

            for i, word in enumerate(ocr_data["text"]):
                if word.strip():  # Exclude empty words
                    x = ocr_data["left"][i]
                    y = ocr_data["top"][i]
                    width = ocr_data["width"][i]
                    height = ocr_data["height"][i]

                    # Calculate coordinates of bounding box
                    x0 = x
                    y0 = y
                    x1 = x + width
                    y1 = y + height

                    print(f"Word: {word}, Coordinates: ({x0}, {y0}, {x1}, {y1})")
            """""
            # Extraire des patterns spécifiques du texte OCR
            reservation_number_matches = findall(self.reservation_number_pattern, ocr_text)
            date_matches = findall(self.date_pattern, ocr_text)
            total_price_matches = findall(self.total_price_pattern, ocr_text)
            plate_number_matches = findall(self.license_plate_pattern, ocr_text, IGNORECASE)
            reservation_number=""
            date_depot=""
            date_restitution=""
            max_total_price=""
            plate_number=""
            found_values=""
            # Garder uniquement la première valeur dans la liste des reservations et ajouter A
            reservation_number = reservation_number_matches[0] if reservation_number_matches else ""
            if reservation_number:
                reservation_number = f'A0{reservation_number}'

            # Valider les dates trouvées au format jj/mm/aaaa et les trier par ordre chronologique par le try/except
            correct_date_matches = []
            for date_match in date_matches:
                try:
                    datetime.strptime(date_match, "%d/%m/%Y")
                    correct_date_matches.append(date_match)
                except ValueError:
                    # Le format de la date n'est pas valide
                    continue
            correct_date_matches = sorted(correct_date_matches, key=lambda x: datetime.strptime(x, "%d/%m/%Y"))

            # Si deux dates sont trouvées, la première chronologiquement est la date de dépôt et la deuxième est la date de restitution
            if len(correct_date_matches) >= 2:
                date_depot = correct_date_matches[0]
                date_restitution = correct_date_matches[1]
            
            # Trouve le prix total
            if total_price_matches:
                max_total_price = max(total_price_matches)

            found_values = [
                value
                for value in search_values
                if value.lower() in ocr_text.lower()
            ]

            if len(plate_number_matches) >= 1:
                plate_number = plate_number_matches[0]
            else:
                plate_number_matches = findall(self.license_plate_pattern1, ocr_text, IGNORECASE)
                if len(plate_number_matches) >= 1:
                    plate_number = plate_number_matches[0]
                else:
                    plate_number_matches = findall(self.license_plate_pattern2, ocr_text, IGNORECASE)
                    if len(plate_number_matches) >=1:
                        plate_number = plate_number_matches[0]
                    else:
                        plate_number_matches = findall(self.license_plate_pattern3, ocr_text, IGNORECASE)
                        if len(plate_number_matches) >=1:
                            plate_number = plate_number_matches[0]
            # Incremente les compteurs si les valeurs sont trouvées

            if reservation_number != "":
                reservation_number_count += 1
            if date_depot != "":
                date_depot_count += 1
            if date_restitution != "":
                date_restitution_count += 1
            if max_total_price != "":
                total_price_count += 1
            if plate_number != "":
                plate_number_count += 1
            if found_values != "":
                apporteur_count += 1

            # Combine le texte reconnu par page en une seule chaîne de caractères avec des sauts de ligne.
            recognized_text = "\n".join([f"Reservation: {reservation_number}", 
                                        f"Date depot: {date_depot}", 
                                        f"Date restitution: {date_restitution}",
                                        f"Montant total: {max_total_price} €",
                                        f"Immatriculation: {plate_number}",
                                        f"Apporteur: {found_values[0]}" if found_values else ""])

            # Enregistre le texte reconnu par page dans un fichier texte
            with open(f"{destination}/page_{page_num+1}_recognized.txt", "w", encoding="utf-8") as f:
                f.write(recognized_text)

            if reservation_number and date_depot and date_restitution and max_total_price and plate_number and found_values != "":
                total_pages_found += 1
                pages_with_patterns.append(page_num + 1)
            row_values = [
                page_num + 1,
                reservation_number,
                date_depot,
                date_restitution,
                max_total_price,
                plate_number,
                found_values[0] if found_values else ""
            ]
            sheet.append(row_values)
        # Ferme le document PDF
        doc.close()

        # Calcul les pourcentages
        reservation_number_percentage = (reservation_number_count / total_pages) * 100
        date_depot_percentage = (date_depot_count / total_pages) * 100
        date_restitution_percentage = (date_restitution_count / total_pages) * 100
        total_price_percentage = (total_price_count / total_pages) * 100
        plate_number_percentage = (plate_number_count / total_pages) * 100
        apporteur_percentage = (apporteur_count / total_pages) * 100

        # Trouver les pages où les patterns ne sont pas trouvés
        pages_with_no_patterns = [
            page_num
            for page_num in range(1, total_pages + 1)
            if page_num not in pages_with_patterns
        ]

        messagebox.showinfo("Succès", "Extraction terminée avec succès !\n\n"
                            f"Pourcentage de Reservation Numbers: {reservation_number_percentage:.2f}%\n"
                            f"Pourcentage de Date Depots: {date_depot_percentage:.2f}%\n"
                            f"Pourcentage de Date Restitutions: {date_restitution_percentage:.2f}%\n"
                            f"Pourcentage de Montant Totals: {total_price_percentage:.2f}%\n"
                            f"Pourcentage de Immatriculations: {plate_number_percentage:.2f}%\n"
                            f"Pourcentage de Apporteurs: {apporteur_percentage:.2f}%\n\n"
                            f"Pages where patterns are not found: {', '.join(map(str, pages_with_no_patterns))}")
        print(f"Pourcentage de Reservation Numbers: {reservation_number_percentage:.2f}%")
        print(f"Pourcentage de Date Depots: {date_depot_percentage:.2f}%")
        print(f"Pourcentage de Date Restitutions: {date_restitution_percentage:.2f}%")
        print(f"Pourcentage de Montant Totals: {total_price_percentage:.2f}%")
        print(f"Pourcentage de Immatriculations: {plate_number_percentage:.2f}%")
        print(f"Pourcentage de Apporteurs: {apporteur_percentage:.2f}%")
        print(pages_with_no_patterns)
        self.pdf_file_path.set("")
        self.destination.set("")
        # Enregistrez le classeur Excel
        excel_filename = f"{destination}/extracted_data_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx"
        workbook.save(excel_filename)
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFFinder(root)
    root.mainloop()
