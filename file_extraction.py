import pypdf
class FileLoader:
    pass

import fitz
import pdfplumber
import pandas as pd

def extract_tables_pdfplumber(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            extracted_table = page.extract_table()
            if extracted_table:
                df = pd.DataFrame(extracted_table[1:], columns=extracted_table[0])  # Convert to DataFrame
                tables.append((page_num, df))

    return tables





def extract_text_pymupdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text")  # Extract text from each page
    return text



def extract_links_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    links = []
    
    for page_num, page in enumerate(doc, start=1):
        for link in page.get_links():
            if "uri" in link:  # Extract only web links
                links.append((page_num, link["uri"]))
    
    return links


def extract_images_from_pdf(pdf_path, output_folder):
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc):
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]

            image_filename = f"{output_folder}/image_{i+1}_{img_index+1}.{image_ext}"
            with open(image_filename, "wb") as img_file:
                img_file.write(image_bytes)
            print(f"Saved: {image_filename}")



def extract_tables_pymupdf(pdf_path):
    doc = fitz.open(pdf_path)
    tables = []
    
    for page_num, page in enumerate(doc, start=1):
        text = page.get_tables("text")  # Extracts raw text (not structured tables)
        tables.append((page_num, text))

    return tables


   

pdf_path = "sample_pdfs/Employee Information Sheet.pdf"  # Replace with your PDF file
output_folder = "images"  # Replace with your output directory
extract_images_from_pdf(pdf_path, output_folder)
print(extract_links_from_pdf(pdf_path))

pdf_text = extract_text_pymupdf(pdf_path)
print(pdf_text)


tables = extract_tables_pdfplumber(pdf_path)

# Print extracted tables
for page, df in tables:
    print(f"Table from Page {page}:\n", df, "\n")