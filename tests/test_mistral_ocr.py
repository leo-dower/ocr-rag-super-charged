
import unittest
from src.ocr.mistral_ocr import MistralOCR

class TestMistralOCR(unittest.TestCase):

    def test_instantiation(self):
        try:
            ocr = MistralOCR(api_key="test_key")
            self.assertIsInstance(ocr, MistralOCR)
        except Exception as e:
            self.fail(f"MistralOCR instantiation failed with {e}")

if __name__ == '__main__':
    unittest.main()
