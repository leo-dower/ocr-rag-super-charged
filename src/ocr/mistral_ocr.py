

import requests
import base64
import uuid
import logging
from typing import Dict
from ..core.base_ocr import BaseOCRProcessor, MIN_TEXT_LENGTH

MISTRAL_OCR_API_URL = "https://api.mistral.ai/v1/ocr"

class MistralOCR(BaseOCRProcessor):
    def __init__(self, api_key=""):
        super().__init__()
        self.api_key = api_key

    def extract_text(self, pdf_path: str, lang: str = 'por') -> str:
        try:
            with open(pdf_path, 'rb') as pdf_file:
                pdf_data = pdf_file.read()
                return self._call_mistral_ocr_api(pdf_data, pdf_path, lang)
        except Exception as e:
            logging.error(f"Erro no processamento com Mistral: {e}")
            return ""

    def _call_mistral_ocr_api(self, pdf_data: bytes, file_name: str, lang: str) -> str:
        base64_pdf = base64.b64encode(pdf_data).decode('utf-8')
        
        lang_mapping = {
            'por': 'portuguese',
            'eng': 'english',
            'spa': 'spanish',
            'fra': 'french',
            'deu': 'german'
        }
        
        payload = {
            "model": "mistral-ocr-latest",
            "id": str(uuid.uuid4()),
            "document": {
                "type": "document_base64",
                "document_base64": base64_pdf,
                "document_name": file_name
            },
            "include_image_base64": False
        }
        
        if lang in lang_mapping:
            payload["language"] = lang_mapping[lang]
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        response = requests.post(
            MISTRAL_OCR_API_URL,
            headers=headers,
            json=payload,
            timeout=120
        )
        
        if response.status_code == 200:
            result = response.json()
            all_text = ""
            if "pages" in result:
                for page in result["pages"]:
                    if "text" in page:
                        all_text += page["text"].strip() + "\n\n"
                    elif "markdown" in page:
                        all_text += page["markdown"].strip() + "\n\n"
            return all_text.strip() if len(all_text.strip()) > MIN_TEXT_LENGTH else ""
        else:
            logging.error(f"Erro na API Mistral OCR: Status {response.status_code} - {response.text}")
            return ""

