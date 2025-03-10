import unittest
from unittest.mock import MagicMock, patch
import os
import json
import pandas as pd
from PIL import Image
from main import PDFLoader,PPTLoader,SQLStorage,FileStorage,DOCXLoader,DataExtractor

class TestFileLoaders(unittest.TestCase):
    def test_pdf_loader(self):
        loader = PDFLoader("test.pdf")
        loader.load_file = MagicMock(return_value="Mock PDF Document")
        self.assertEqual(loader.load_file(), "Mock PDF Document")
    
    def test_docx_loader(self):
        loader = DOCXLoader("test.docx")
        loader.load_file = MagicMock(return_value="Mock DOCX Document")
        self.assertEqual(loader.load_file(), "Mock DOCX Document")
    
    def test_ppt_loader(self):
        loader = PPTLoader("test.pptx")
        loader.load_file = MagicMock(return_value="Mock PPT Document")
        self.assertEqual(loader.load_file(), "Mock PPT Document")

class TestDataExtractor(unittest.TestCase):
    def setUp(self):
        self.mock_loader = MagicMock()
        self.extractor = DataExtractor(self.mock_loader)
    
    def test_extract_text(self):
        self.mock_loader.load_file.return_value = "Mock Text Content"
        self.extractor.extract_text = MagicMock(return_value=[{"text": ["Sample Text"]}])
        result = self.extractor.extract_text()
        self.assertEqual(result, [{"text": ["Sample Text"]}])
    
    def test_extract_links(self):
        self.extractor.extract_links = MagicMock(return_value=["http://example.com"])
        result = self.extractor.extract_links()
        self.assertIn("http://example.com", result)
    
    def test_extract_images(self):
        self.extractor.extract_images = MagicMock(return_value=[{"filename": "image1.jpg", "width": 800, "height": 600}])
        result = self.extractor.extract_images()
        self.assertEqual(result[0]["filename"], "image1.jpg")
    
    def test_extract_tables(self):
        df = pd.DataFrame({"Column1": ["Data1"], "Column2": ["Data2"]})
        self.extractor.extract_tables = MagicMock(return_value=[df])
        result = self.extractor.extract_tables()
        self.assertEqual(result[0].iloc[0, 0], "Data1")

class TestStorage(unittest.TestCase):
    def setUp(self):
        self.mock_extractor = MagicMock()
        self.mock_extractor.file_path = "test.pdf"
        self.mock_extractor.extract_text.return_value = [{"text": ["Sample Text"]}]
        self.mock_extractor.extract_links.return_value = ["http://example.com"]
        self.mock_extractor.extract_images.return_value = [{"filename": "image1.jpg", "width": 800, "height": 600}]
        self.mock_extractor.extract_tables.return_value = [pd.DataFrame({"Column1": ["Data1"]})]
    
    @patch("builtins.open", new_callable=unittest.mock.mock_open)
    def test_file_storage(self, mock_open):
        file_storage = FileStorage()
        file_storage.store(self.mock_extractor)
        self.assertTrue(mock_open.called)
    
    @patch("mysql.connector.connect")
    def test_sql_storage(self, mock_mysql):
        mock_conn = mock_mysql.return_value
        mock_cursor = mock_conn.cursor.return_value
        sql_storage = SQLStorage()
        sql_storage.store(self.mock_extractor)
        self.assertTrue(mock_cursor.execute.called)
        self.assertTrue(mock_conn.commit.called)

if __name__ == "__main__":
    unittest.main()
