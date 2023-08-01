import tkinter as tk
from tkinter import filedialog, messagebox
from os import path
from threading import Thread
import fitz
from PIL import Image
import pytesseract
import re
from tkinter import ttk
from datetime import datetime

class PDFFinder:
    def __init__(self, root):
        self.root = root
        self.root.title("Extract PDF Data")
        self.root.geometry("300x300")

        self.pdf_file_path = tk.StringVar()
        self.destination = tk.StringVar()
        self.reservation_number_pattern = r"(\d{6})"
        self.date_pattern = r"(\d{1,2}/\d{1,2}/\d{4})"
        self.total_price_pattern = r"(\d+,\d{2})"
        self.plate_number_pattern =  r"\b(?:[A-Z]{2}\s?\d{3}\s?[A-Z]{2}|[A-Z]{2}\d{2}\s?[A-Z]{2}|[A-Z]{2}\s?\d{3}[A-Z]{2})\b"
       
        self.create_widgets()

    def create_widgets(self):
        # Create and place widgets here...
        self.label_pdf_file = tk.Label(self.root, text="PDF File:")
        self.label_pdf_file.pack(pady=5)

        self.entry_pdf_file = tk.Entry(self.root, textvariable=self.pdf_file_path)
        self.entry_pdf_file.pack(pady=5)

        self.btn_browse_pdf = tk.Button(self.root, text="Browse", command=self.browse_pdf)
        self.btn_browse_pdf.pack(pady=5)

        self.label_destination = tk.Label(self.root, text="Destination:")
        self.label_destination.pack(pady=5)

        self.entry_destination = tk.Entry(self.root, textvariable=self.destination)
        self.entry_destination.pack(pady=5)

        self.btn_browse_destination = tk.Button(self.root, text="Browse", command=self.browse_destination)
        self.btn_browse_destination.pack(pady=5)

        self.btn_start_extraction = tk.Button(self.root, text="Start Extraction", command=self.start_extraction)
        self.btn_start_extraction.pack(pady=10)

        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)

    def browse_pdf(self):
        if file_path:= filedialog.askopenfilename(
            filetypes=[("PDF Files", 
                        "*.pdf")]
        ):
            self.pdf_file_path.set(file_path)

    def browse_destination(self):
        if folder_path:= filedialog.askdirectory():
            self.destination.set(folder_path)

    def validate_pdf_file_path(self, pdf_file_path):
        # Validate the PDF file path
        if not path.isfile(pdf_file_path):
            messagebox.showerror("Error", "Invalid PDF file path.")
            return False
        elif not pdf_file_path.lower().endswith(".pdf"):
            messagebox.showerror("Error", "Please select a PDF file.")
            return False
        return True

    def start_extraction(self):
        pdf_file_path = self.pdf_file_path.get()
        destination = self.destination.get()

        if not destination:
            messagebox.showerror("Error", "Please provide a destination folder.")
            return

        if self.validate_pdf_file_path(pdf_file_path):
            thread = Thread(target=self.process_pdf, args=(pdf_file_path, destination))
            thread.start()

    def process_pdf(self, pdf_file_path, destination):
        doc = fitz.open(pdf_file_path)
        total_pages = len(doc)

        # Define the values to search for
        search_values = ["parkcloud", "travelcar", "zenpark", "parkos", "travelercar"]

        # Counters for patterns and values
        total_pages_found = 0
        reservation_number_count = 0
        date_depot_count = 0
        date_restitution_count = 0
        total_price_count = 0
        plate_number_count = 0
        apporteur_count = 0

        # List to store pages where patterns are found
        pages_with_patterns = []

        for page_num in range(total_pages):
            # Update progress bar
            progress_percent = (page_num + 1) / total_pages * 100
            self.progress_var.set(progress_percent)
            self.root.update()

            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Increase DPI
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Perform OCR using pytesseract
            ocr_text = pytesseract.image_to_string(image)

            # Extract specific patterns from OCR text
            reservation_number_matches = re.findall(self.reservation_number_pattern, ocr_text)
            date_matches = re.findall(self.date_pattern, ocr_text)
            total_price_matches = re.findall(self.total_price_pattern, ocr_text)
            plate_number_matches = re.findall(self.plate_number_pattern, ocr_text, re.IGNORECASE)

            # Keep only the first match for reservation number and add 'A' at the beginning
            reservation_number = reservation_number_matches[0] if reservation_number_matches else ""
            if reservation_number:
                reservation_number = f'A{reservation_number}'

            # Validate the date format as "dd/mm/yyyy" using try-except
            correct_date_matches = []
            for date_match in date_matches:
                try:
                    datetime.strptime(date_match, "%d/%m/%Y")
                    correct_date_matches.append(date_match)
                except ValueError:
                    # Date format does not match
                    continue
            correct_date_matches = sorted(correct_date_matches, key=lambda x: datetime.strptime(x, "%d/%m/%Y"))

            # If there are at least two dates, identify "Date depot" and "Date restitution"
            date_depot = ""
            date_restitution = ""
            if len(correct_date_matches) >= 2:
                date_depot = correct_date_matches[0]
                date_restitution = correct_date_matches[1]
            
            # Find the highest total price
            max_total_price = ""
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
                plate_number = ""

            # Count patterns and values
            if reservation_number:
                reservation_number_count += 1
            if date_depot:
                date_depot_count += 1
            if date_restitution:
                date_restitution_count += 1
            if max_total_price:
                total_price_count += 1
            if plate_number:
                plate_number_count += 1
            if found_values:
                apporteur_count += 1

            # Combine the recognized text from all patterns and the found values
            recognized_text = "\n".join([f"Reservation: {reservation_number}", 
                                        f"Date depot: {date_depot}", 
                                        f"Date restitution: {date_restitution}",
                                        f"Montant total: {max_total_price} €",
                                        f"Immatriculation: {plate_number}",
                                        f"Apporteur: {found_values[0]}" if found_values else ""])

            # Store the extracted recognized text per page
            #with open(f"{destination}/page_{page_num+1}_recognized.txt", "w", encoding="utf-8") as f:
            #    f.write(recognized_text)

            if reservation_number and date_depot and date_restitution and max_total_price and plate_number and found_values != "":
                total_pages_found += 1
                pages_with_patterns.append(page_num + 1)  # Add 1 to convert to 1-based page numbering
            else:
                print(ocr_text)
                            # Store the extracted recognized text per page
                with open(f"{destination}/page_{page_num+1}_recognized.txt", "w", encoding="utf-8") as f:
                    f.write(recognized_text)
        # Close the PDF document
        doc.close()

        # Calculate percentages
        reservation_number_percentage = (reservation_number_count / total_pages) * 100
        date_depot_percentage = (date_depot_count / total_pages) * 100
        date_restitution_percentage = (date_restitution_count / total_pages) * 100
        total_price_percentage = (total_price_count / total_pages) * 100
        plate_number_percentage = (plate_number_count / total_pages) * 100
        apporteur_percentage = (apporteur_count / total_pages) * 100

        # Calculate pages where patterns are not found
        pages_with_no_patterns = [
            page_num
            for page_num in range(1, total_pages + 1)  # 1-based page numbering
            if page_num not in pages_with_patterns
        ]

        messagebox.showinfo("Success", "Recognized text extracted successfully!\n\n"
                            f"Percentage of Reservation Numbers: {reservation_number_percentage:.2f}%\n"
                            f"Percentage of Date Depots: {date_depot_percentage:.2f}%\n"
                            f"Percentage of Date Restitutions: {date_restitution_percentage:.2f}%\n"
                            f"Percentage of Montant Totals: {total_price_percentage:.2f}%\n"
                            f"Percentage of Immatriculations: {plate_number_percentage:.2f}%\n"
                            f"Percentage of Apporteurs: {apporteur_percentage:.2f}%\n\n"
                            f"Pages where patterns are not found: {', '.join(map(str, pages_with_no_patterns))}")
        print(f"Percentage of Reservation Numbers: {reservation_number_percentage:.2f}%")
        print(f"Percentage of Date Depots: {date_depot_percentage:.2f}%")
        print(f"Percentage of Date Restitutions: {date_restitution_percentage:.2f}%")
        print(f"Percentage of Montant Totals: {total_price_percentage:.2f}%")
        print(f"Percentage of Immatriculations: {plate_number_percentage:.2f}%")
        print(f"Percentage of Apporteurs: {apporteur_percentage:.2f}%")
        print(pages_with_no_patterns)
        self.pdf_file_path.set("")
        self.destination.set("")
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFFinder(root)
    root.mainloop()
