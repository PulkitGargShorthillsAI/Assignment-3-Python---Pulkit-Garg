
# Assignment-3 (Python) - Pulkit Garg

## Overview

This project provides a modular Python class structure to extract **text, hyperlinks, images, and tables** from various document formats (**PDF, DOCX, PPTX**) while capturing metadata. It follows an **object-oriented approach** using abstract classes for flexibility and concrete classes for specific implementations.

## Features

- **Extract Text**: Retrieves text content along with metadata (page numbers, font styles, headings).
- **Extract Hyperlinks**: Extracts hyperlinks along with metadata (linked text, URL, paragraph/slide number).
- **Extract Images**: Extracts images while preserving metadata (resolution, format, page number/slide number).
- **Extract Tables**: Extracts tabular data while preserving metadata (dimensions, page/slide number).
- **Storage Support**: Saves extracted data to files and MySQL databases.

## Class Structure

### File Loaders
Abstract class: `FileLoader`
- **PDFLoader**: Loads and processes PDF files using `PyMuPDF (fitz)`.
- **DOCXLoader**: Loads and processes DOCX files using `python-docx`.
- **PPTLoader**: Loads and processes PPTX files using `python-pptx`.

### Data Extraction
Class: `DataExtractor`
- **extract_text()**: Extracts text with metadata (font size, bold, italic, heading detection).
- **extract_links()**: Extracts hyperlinks with metadata.
- **extract_images()**: Extracts images with resolution and metadata.
- **extract_tables()**: Extracts tabular data and preserves structure.

### Storage Classes
Abstract class: `Storage`
- **FileStorage**: Stores extracted text, links, images, and tables as separate files.
- **SQLStorage**: Stores extracted text, links, images, and tables in a MySQL database.

## Technologies Used

- **Python 3.x**
- **PyMuPDF (fitz)** - PDF processing
- **python-docx** - DOCX processing
- **python-pptx** - PPTX processing
- **pdfplumber** - Table extraction from PDFs
- **Pandas** - Data handling
- **Pillow (PIL)** - Image processing
- **MySQL Connector** - Database storage

---

## Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/your-repo.git
cd your-repo

### 2. Install Dependencies
Install required Python packages:
```bash
pip install -r requirements.txt
```

### 3. Set Up MySQL Database (Optional)
If using MySQL storage, create a database and update `SQLStorage` class in `main.py`:
```sql
CREATE DATABASE assignment3;
```
Ensure the MySQL user credentials are correct:
```python
self.conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="yourpassword",
    database="assignment3"
)
```

---

## Usage

### Extract Data from a Document
Modify the `file_path` in `main.py`:
```python
file_path = "assets/sample_pdfs/test1.pdf"  # or .docx / .pptx
```
Run the script:
```bash
python main.py
```

### Running Unit Tests
Unit tests are included in `testing.py` to validate functionality.
Run all test cases:
```bash
python -m unittest testing.py
```

---

## Unit Testing

The test cases in `testing.py` validate the functionality of the program:

| TEST CASE ID | SECTION     | SUB-SECTION  | TEST CASE TITLE           | TEST DESCRIPTION                          | PRECONDITIONS     | TEST DATA             | TEST STEPS                         | EXPECTED RESULT                     | ACTUAL RESULT | STATUS |
|-------------|------------|-------------|---------------------------|------------------------------------------|------------------|-----------------------|-------------------------------------|-------------------------------------|--------------|--------|
| TC_01       | File Loading | PDF        | Verify PDF Loading       | Check if PDFLoader loads PDF correctly  | Valid PDF file   | sample.pdf            | Call `PDFLoader(file_path)`       | Returns a valid PDF file object     | As expected  | PASS   |
| TC_02       | File Loading | DOCX       | Verify DOCX Loading      | Check if DOCXLoader loads DOCX correctly | Valid DOCX file  | sample.docx           | Call `DOCXLoader(file_path)`      | Returns a valid DOCX file object    | As expected  | PASS   |
| TC_03       | File Loading | PPTX       | Verify PPTX Loading      | Check if PPTLoader loads PPTX correctly | Valid PPTX file  | sample.pptx           | Call `PPTLoader(file_path)`       | Returns a valid PPTX file object    | As expected  | PASS   |
| TC_04       | Data Extraction | Text    | Extract Text from PDF    | Extract text from a PDF document       | Valid PDF file   | sample.pdf            | Call `extract_text()`             | Returns extracted text & metadata  | As expected  | PASS   |
| TC_05       | Data Extraction | Links   | Extract Links from DOCX  | Extract links from a DOCX document     | Valid DOCX file  | sample.docx           | Call `extract_links()`            | Returns extracted links & metadata | As expected  | PASS   |
| TC_06       | Data Extraction | Images  | Extract Images from PPTX | Extract images from a PPTX document    | Valid PPTX file  | sample.pptx           | Call `extract_images()`           | Returns image metadata             | As expected  | PASS   |
| TC_07       | Data Extraction | Tables  | Extract Tables from DOCX | Extract tables from a DOCX document    | Valid DOCX file  | sample.docx           | Call `extract_tables()`           | Returns table metadata             | As expected  | PASS   |
| TC_08       | Storage      | File       | Store Extracted Data in Files | Verify file-based storage | Valid extracted data | Extracted text, links, images, tables | Call `FileStorage().store()` | Stores extracted data in files | As expected | PASS |
| TC_09       | Storage      | Database   | Store Extracted Data in MySQL | Verify MySQL storage | Valid extracted data | Extracted text, links, images, tables | Call `SQLStorage().store()` | Stores extracted data in MySQL | As expected | PASS |

---

## Expected Outputs

### Text Extraction Output (Example)
```json
{
  "page_number": 1,
  "text": ["Extracted text from document..."],
  "metadata": [
    {"text": "Extracted text", "font_size": 12, "bold": false, "italic": false, "font_style": "Times New Roman"}
  ]
}
```

### Link Extraction Output (Example)
```json
[
  {"paragraph_number": 2, "text": "Click here", "url": "https://example.com"}
]
```

### Image Extraction Output (Example)
```json
[
  {"filename": "image_1.png", "page_number": 1, "width": 800, "height": 600}
]
```

---

## Contribution Guidelines

- Fork the repository.
- Create a feature branch.
- Commit your changes.
- Push to your branch and create a pull request.

---

## Screenshots
  - ![alt text](<out_put_images/text_with_metadata.png>)

  - ![alt text](<out_put_images/extracted_data_stored_in_sql_db.png>)

  - ![alt text](<out_put_images/unit_testing_successful.png>)

---

## Contact

For questions, open an issue or contact the project maintainer.

---

Happy coding! 🚀
```

This `README.md` covers:
- **Project Overview**
- **Features**
- **Class Structure**
- **Installation & Usage**
- **Unit Testing**
- **Expected Outputs**
- **Contribution Guidelines**
- **License & Contact Information**
