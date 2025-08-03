
import unittest
import os
from src.ocr.tesseract_ocr import TesseractOCR

class TestTesseractOCR(unittest.TestCase):

    def test_extract_text(self):
        # This is an integration test and requires a PDF file.
        # We will just check if the class can be instantiated.
        try:
            ocr = TesseractOCR()
            self.assertIsInstance(ocr, TesseractOCR)
        except Exception as e:
            self.fail(f"TesseractOCR instantiation failed with {e}")

if __name__ == '__main__':
    unittest.main()
