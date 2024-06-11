import os
import fitz  # PyMuPDF
from PyPDF2 import PdfFileReader
import pandas as pd

def extract_text_info(pdf_path):
    try:
        document = fitz.open(pdf_path)
        first_page = document[0]
        text = first_page.get_text()

        # Basic extraction (needs to be improved based on the PDF layout)
        title = text.split('\n')[0]
        author = "Unknown"
        year = "Unknown"

        for line in text.split('\n'):
            if "author" in line.lower():
                author = line.split(":")[-1].strip()
            if "year" in line.lower():
                year = line.split(":")[-1].strip()
        
        return title, author, year
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None, None, None

def extract_metadata(pdf_path):
    try:
        with open(pdf_path, 'rb') as f:
            reader = PdfFileReader(f)
            info = reader.getDocumentInfo()
            title = info.title if info.title else "Unknown"
            author = info.author if info.author else "Unknown"
            year = "Unknown"  # Year is typically not in metadata
            return title, author, year
    except Exception as e:
        print(f"Error reading metadata from {pdf_path}: {e}")
        return None, None, None

def extract_info_from_folders(base_folder):
    data = []
    for root, dirs, files in os.walk(base_folder):
        for file in files:
            if file.endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                title, author, year = extract_metadata(pdf_path)
                if title == "Unknown" and author == "Unknown":
                    title, author, year = extract_text_info(pdf_path)
                data.append([pdf_path, title, author, year])
                print(f"File: {pdf_path}")
                print(f"Title: {title}, Author: {author}, Year: {year}")
                print("-" * 40)
    return data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=["File Path", "Title", "Author", "Year"])
    df.to_excel(output_file, index=False)

def save_to_txt(data, output_file):
    with open(output_file, 'w') as f:
        for item in data:
            f.write(f"File: {item[0]}\nTitle: {item[1]}\nAuthor: {item[2]}\nYear: {item[3]}\n{'-'*40}\n")

base_folder = r"E:\Thesis - Writing Selva\LR\Normal"  # Change this to your folder path
data = extract_info_from_folders(base_folder)
save_to_excel(data, "output.xlsx")
save_to_txt(data, "output.txt")

