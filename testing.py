import unittest
import os
from unittest.mock import MagicMock, patch
from main import PDFLoader, DOCXLoader, PPTLoader, DataExtractor, FileStorage, SQLStorage

class TestFileExtraction(unittest.TestCase):
    def setUp(self):
        """Setup test files (Ensure these exist in the test directory)"""
        self.pdf_path = "assets/sample_pdfs/test1.pdf"
        self.docx_path = "assets/sample_docx/demo.docx"
        self.pptx_path = "assets/sample_pptx/ppt_test.pptx"
        
    def test_pdf_loader(self):
        """Test PDF file loading"""
        loader = PDFLoader(self.pdf_path)
        self.assertTrue(loader.load_file())

    def test_docx_loader(self):
        """Test DOCX file loading"""
        loader = DOCXLoader(self.docx_path)
        self.assertTrue(loader.load_file())
    
    def test_pptx_loader(self):
        """Test PPTX file loading"""
        loader = PPTLoader(self.pptx_path)
        self.assertTrue(loader.load_file())

    @patch("main.fitz.open")
    def test_extract_text_pdf(self, mock_fitz):
        """Test text extraction from a PDF"""
        mock_fitz.return_value = MagicMock()
        loader = PDFLoader(self.pdf_path)
        extractor = DataExtractor(loader)
        self.assertIsInstance(extractor.extract_text(), list)
    
    @patch("main.docx.Document")
    def test_extract_text_docx(self, mock_docx):
        """Test text extraction from a DOCX"""
        mock_docx.return_value = MagicMock()
        loader = DOCXLoader(self.docx_path)
        extractor = DataExtractor(loader)
        self.assertIsInstance(extractor.extract_text(), list)

    @patch("main.pptx.Presentation")
    def test_extract_text_pptx(self, mock_pptx):
        """Test text extraction from a PPTX"""
        mock_pptx.return_value = MagicMock()
        loader = PPTLoader(self.pptx_path)
        extractor = DataExtractor(loader)
        self.assertIsInstance(extractor.extract_text(), list)
    
    def test_extract_links(self):
        """Test extracting links from files"""
        loader = DOCXLoader(self.docx_path)
        extractor = DataExtractor(loader)
        self.assertIsInstance(extractor.extract_links(), list)
    
    def test_extract_images(self):
        """Test extracting images from files"""
        loader = PDFLoader(self.pdf_path)
        extractor = DataExtractor(loader)
        self.assertIsInstance(extractor.extract_images(), list)
    
    def test_extract_tables(self):
        """Test extracting tables from files"""
        loader = DOCXLoader(self.docx_path)
        extractor = DataExtractor(loader)
        self.assertIsInstance(extractor.extract_tables(), list)
    
    @patch("main.mysql.connector.connect")
    def test_sql_storage(self, mock_db):
        """Test storing extracted data in MySQL"""
        mock_db.return_value.cursor.return_value.execute = MagicMock()
        loader = PDFLoader(self.pdf_path)
        extractor = DataExtractor(loader)
        storage = SQLStorage()
        storage.store(extractor)
        mock_db.return_value.commit.assert_called()
    
    def test_file_storage(self):
        """Test storing extracted data in files"""
        loader = DOCXLoader(self.docx_path)
        extractor = DataExtractor(loader)
        storage = FileStorage()
        storage.store(extractor)
        self.assertTrue(os.path.exists("extracted_text.txt"))

if __name__ == "__main__":
    unittest.main()
