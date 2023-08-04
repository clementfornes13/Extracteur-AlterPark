import tkinter as tk
from tkinter import filedialog, messagebox
from os import path
from threading import Thread
import fitz
from PIL import Image
import pytesseract
from re import findall, IGNORECASE
from tkinter import ttk
from datetime import datetime
from openpyxl import Workbook
from time import perf_counter
import webbrowser

class PDFFinder:
    def __init__(self, root):
        
        self.root = root
        self.root.title("Extraire Data Alter Park")
        self.root.geometry("300x330")
        self.root.resizable(False, False)

        self.reservation_number_pattern = r"(?<![0-9])0*(\d{5})"
        self.date_pattern = r"(\d{1,2}/\d{1,2}/\d{4})"
        self.total_price_pattern = r"(\d+,\d{2})"
        self.license_plate_pattern =  r"\b(?:[A-Z]{2}\s?\d{3}\s?[A-Z]{2}|[A-Z]{2}\d{2}\s?[A-Z]{2}|[A-Z]{2}\s?\d{3}[A-Z]{2})\b"
        self.license_plate_pattern1 = r'\b[A-Z]{2}\s?[-]?\s?\d{3}\s?[-]?\s?[A-Z]{2}\b|\b\d{3}\s?[-]?\s?[A-Z]{2}\s?[-]?\s?\d{3}\b'
        self.license_plate_pattern2 = r'\b[A-Z]{2}\d{2}\s?[A-Z]{3}\b|\b[A-Z]{2}\d{3}\s?[A-Z]{3}\b'
        self.license_plate_pattern3 = r'\b[A-Z]{3}\s?[-]?\s?\d{4}\b|\b\d{4}\s?[-]?\s?[A-Z]{3}\b'
        self.license_plate_pattern4 = r'\d{3}[A-Z]{3}\d{2}'
        self.license_plate_pattern5 = r'\d{4}[A-Z]{2}\d{2}'
        self.pdf_file_path = tk.StringVar()
        self.destination = tk.StringVar()
        
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
        if not path.isfile(pdf_file_path):
            messagebox.showerror("Erreur", "Chemin du fichier PDF invalide.")
            return False
        elif not pdf_file_path.lower().endswith(".pdf"):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier PDF.")
            return False
        return True
    
    def validate_destination(self, destination):
        if destination == "":
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier de destination.")
            return False
        elif not path.isdir(destination):
            messagebox.showerror("Erreur", "Chemin du dossier de destination invalide.")
            return False
        return True
        
    def start_extraction(self):
        pdf_file_path = self.pdf_file_path.get()
        destination = self.destination.get()
        
        if self.validate_pdf_file_path(pdf_file_path) and self.validate_destination(destination):
            thread = Thread(target=self.process_pdf, args=(pdf_file_path, destination))
            thread.start()
    
    def update_progress(self, page_num, total_pages):
        # Mettre à jour la barre de progression et le pourcentage de progression
        progress_percent = (page_num + 1) / total_pages * 100
        self.progress_var.set(progress_percent)
        self.label_progress.config(text=f"{progress_percent:.2f}%")
        # Calculer le temps restant estimé
        time_remaining = (perf_counter() - self.start_time) / (page_num + 1) * (total_pages - page_num - 1)
        self.label_time_remaining.config(text=f"Temps restant : {time_remaining:.0f} secondes")
        self.root.update()
        
    def show_success_message(self, pages_with_no_patterns, excel_filename):
        message = (
            f"Extraction terminée avec succès!\n\n"
            f"Pages contenant des valeurs non trouvées : {', '.join(map(str, pages_with_no_patterns))}\n"
            f"Classeur Excel sauvegardé dans {excel_filename}"
        )
        messagebox.showinfo("Succès", message)
    
    def initialize_excel(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data Extraite"
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
        return sheet, workbook
    
    def initialize_values(self):
        search_values = ["parkcloud", "travelcar", "zenpark", "parkos", "travelercar"]
        dpi_factor = 400
        total_pages_found = 0
        reservation_number_count = 0
        date_depot_count = 0
        date_restitution_count = 0
        total_price_count = 0
        plate_number_count = 0
        apporteur_count = 0
        pages_with_patterns = []
        return search_values, dpi_factor, total_pages_found, reservation_number_count, date_depot_count, date_restitution_count, total_price_count, plate_number_count, apporteur_count, pages_with_patterns
    
    def process_pdf(self, pdf_file_path, destination):
        doc = fitz.open(pdf_file_path)
        total_pages = len(doc)
        search_values, dpi_factor, total_pages_found, reservation_number_count, date_depot_count, date_restitution_count, total_price_count, plate_number_count, apporteur_count, pages_with_patterns = self.initialize_values()
        sheet, workbook = self.initialize_excel()
        self.start_time = perf_counter()
        for page_num in range(total_pages):
            self.update_progress(page_num, total_pages)
            ocr_text = self.launch_ocr(doc, page_num, dpi_factor)
            reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values = self.process_ocr(ocr_text, search_values)
            self.increase_count(reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values, reservation_number_count, date_depot_count, date_restitution_count, total_price_count, plate_number_count, apporteur_count)
            self.log_page(destination, ocr_text, page_num, reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values)
            self.excel_add(sheet, page_num, reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values, pages_with_patterns, total_pages_found)
        pages_with_no_patterns = self.finish_extraction(doc, total_pages, pages_with_patterns)
        excel_filename = self.excel_save_file(destination, workbook)
        self.show_success_message(pages_with_no_patterns, excel_filename)
        
    def finish_extraction(self, doc, total_pages, pages_with_patterns):
        doc.close()
        pages_with_no_patterns = [
            page_num
            for page_num in range(1, total_pages + 1)
            if page_num not in pages_with_patterns
        ]
        self.pdf_file_path.set("")
        self.destination.set("")
        return pages_with_no_patterns
    
    def launch_ocr(self, doc, page_num, dpi_factor):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_factor / 72, dpi_factor / 72))
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ocr_text = pytesseract.image_to_string(image)
        ocr_text = ocr_text.replace('I', '1').replace('O', '0').replace('U', 'V')
        ocr_text = ocr_text.replace('parkcl0vd', 'parkcloud').replace('park0s', 'parkos').replace('PARKCL0VD','PARKCLOUD').replace('PARK0S','PARKOS')
        return ocr_text
    
    def process_ocr(self, ocr_text, search_values):
        if 'ASE' not in ocr_text:
            reservation_number_matches = findall(self.reservation_number_pattern, ocr_text)
        else:
            reservation_number_matches = []
        date_matches = findall(self.date_pattern, ocr_text)
        total_price_matches = findall(self.total_price_pattern, ocr_text)
        plate_number_matches = findall(self.license_plate_pattern, ocr_text, IGNORECASE)
        reservation_number=""
        date_depot=""
        date_restitution=""
        plate_number=""
        found_values=""
        reservation_number = reservation_number_matches[0] if reservation_number_matches else ""
        if reservation_number:
            reservation_number = f'A0{reservation_number}'
        correct_date_matches = []
        for date_match in date_matches:
            try:
                datetime.strptime(date_match, "%d/%m/%Y")
                correct_date_matches.append(date_match)
            except ValueError:
                continue
        correct_date_matches = sorted(correct_date_matches, key=lambda x: datetime.strptime(x, "%d/%m/%Y"))
        if len(correct_date_matches) >= 2:
            date_depot = correct_date_matches[0]
            date_restitution = correct_date_matches[1]
        max_total_price = max(total_price_matches) if total_price_matches else ""
        found_values = [
            value
            for value in search_values
            if value.lower() in ocr_text.lower()
        ]
        plate_number = self.license_plate_check(plate_number_matches, ocr_text)
        return reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values
    
    def license_plate_check(self, plate_number_matches, ocr_text):
        patterns = [
            self.license_plate_pattern1,
            self.license_plate_pattern2,
            self.license_plate_pattern3,
            self.license_plate_pattern4,
            self.license_plate_pattern5
        ]
        for pattern in patterns:
            plate_number_matches = findall(pattern, ocr_text, IGNORECASE)
            if len(plate_number_matches) >= 1:
                return plate_number_matches[0]
        
        return None
    
    def increase_count(self, reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values, reservation_number_count, date_depot_count, date_restitution_count, total_price_count, plate_number_count, apporteur_count):
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
        return reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values, reservation_number_count, date_depot_count, date_restitution_count, total_price_count, plate_number_count, apporteur_count
    
    def excel_add(self, sheet, page_num, reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values, pages_with_patterns, total_pages_found):
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

    def log_page(self, destination, ocr_text, page_num, reservation_number, date_depot, date_restitution, max_total_price, plate_number, found_values):
        recognized_text = "\n".join([f"Reservation: {reservation_number}", 
                                    f"Date depot: {date_depot}", 
                                    f"Date restitution: {date_restitution}",
                                    f"Montant total: {max_total_price} €",
                                    f"Immatriculation: {plate_number}",
                                    f"Apporteur: {found_values[0]}" if found_values else ""])
        with open(f"{destination}/page_{page_num+1}_ocr.txt", "w", encoding="utf-8") as f:
            f.write(ocr_text)
        with open(f"{destination}/page_{page_num+1}_recognized.txt", "w", encoding="utf-8") as f:
            f.write(recognized_text)
            
    def excel_save_file(self, destination, workbook):
        excel_filename = f"{destination}/extracted_data_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx"
        workbook.save(excel_filename)
        return excel_filename

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFFinder(root)
    root.mainloop()
