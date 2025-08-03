

import pytesseract
from PIL import Image, ImageOps
from pdf2image import convert_from_path
from bs4 import BeautifulSoup
import logging
from typing import List
from ..core.base_ocr import BaseOCRProcessor, MIN_TEXT_LENGTH

class TesseractOCR(BaseOCRProcessor):
    def __init__(self, poppler_path=None):
        super().__init__()
        self.poppler_path = poppler_path

    def _preprocess_image(self, image: Image.Image) -> Image.Image:
        return ImageOps.autocontrast(image.convert('L').point(lambda x: 0 if x < 128 else 255))

    def extract_text(self, pdf_path: str, lang: str = 'por') -> str:
        try:
            images = convert_from_path(pdf_path, poppler_path=self.poppler_path)
            return self._perform_ocr(images, lang)
        except Exception as e:
            logging.error(f"Erro no processamento com Tesseract: {e}")
            return ""

    def _perform_ocr(self, images: List[Image.Image], lang: str) -> str:
        text = ""
        for image in images:
            if self.stop_event.is_set():
                break

            processed = self._preprocess_image(image)
            hocr_data = pytesseract.image_to_pdf_or_hocr(
                processed,
                extension='hocr',
                config=f'--psm 1 -l {lang}'
            )
            
            soup = BeautifulSoup(hocr_data, 'html.parser')
            paragraphs = soup.find_all('p', class_='ocr_par')
            
            for para in paragraphs:
                lines = []
                for line in para.find_all('span', class_='ocr_line'):
                    words = line.find_all('span', class_='ocrx_word')
                    line_text = ' '.join(self._process_words(words))
                    lines.append(line_text)
                text += '\n'.join(lines) + '\n\n'

        return text if len(text.strip()) > MIN_TEXT_LENGTH else ""

    def _process_words(self, words: List[BeautifulSoup]) -> List[str]:
        processed_words = []
        for word in words:
            word_text = word.get_text().strip()
            if 'bold' in word.get('class', []):
                processed_words.append(f"**{word_text}**")
            else:
                processed_words.append(word_text)
        return processed_words
