

import threading
from typing import List, Tuple

MIN_TEXT_LENGTH = 50

class BaseOCRProcessor:
    """Classe base para processadores de OCR"""
    
    def __init__(self):
        self.stop_event = threading.Event()
    
    def extract_text(self, pdf_path: str, lang: str = 'por') -> str:
        """Método a ser implementado pelas subclasses"""
        raise NotImplementedError("Método deve ser implementado pela subclasse")
    
    def get_paragraphs(self, text: str) -> List[Tuple[str, str]]:
        """
        Divide o texto em parágrafos preservando estrutura básica
        Retorna uma lista de tuplas (texto, tipo_paragrafo)
        """
        # Padrões para identificação de tipos de parágrafos
        artigo_pattern = re.compile(r'(artigo|art\.?)\s*\d+º?', re.IGNORECASE)
        titulo_pattern = re.compile(r'^(TÍTULO|CAPÍTULO|SEÇÃO)\s+[IVXLCDM0-9]+', re.IGNORECASE)
        
        # Dividir o texto em parágrafos potenciais
        raw_paragraphs = [p.strip() for p in re.split(r'\n{2,}', text) 
                         if p.strip() and any(c.isalnum() for c in p) 
                         and len(p.strip()) > 10]
        
        processed_paragraphs = []
        
        for para in raw_paragraphs:
            # Substituir quebras de linha únicas por espaços para melhorar a leitura
            clean_para = re.sub(r'(?<!\n)\n(?!\n)', ' ', para).strip()
            
            # Identificar o tipo de parágrafo
            if titulo_pattern.search(clean_para):
                para_type = "titulo"
            elif artigo_pattern.search(clean_para):
                para_type = "artigo"
            elif "**" in clean_para:
                para_type = "destaque"
            else:
                para_type = "normal"
                
            processed_paragraphs.append((clean_para, para_type))
            
        return processed_paragraphs

