

import re
import pandas as pd
from datetime import datetime
from typing import Dict, List, Any, Optional
import json
from mistralai import Mistral
import logging
from pdfminer.high_level import extract_text

class DocumentDataExtractor:
    """
    Extrator avançado de dados estruturados para diferentes tipos de documentos
    """
    def __init__(self, mistral_api_key: Optional[str] = None):
        self.mistral_client = Mistral(api_key=mistral_api_key) if mistral_api_key else None
        self.extraction_patterns = {
            'juridico': {
                'tipo_documento': r'\b(PROCESSO|PETIÇÃO|RECURSO|AÇÃO)\b',
                'numero_processo': r'\b(?:PROCESSO|PROTOCOLO)\s*[Nº]?\s*(\d{4,20})\b',
                'data_documento': r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b',
                'valor_causa': r'\bVALOR\s*(?:DA\s*CAUSA)?\s*[R$]?\s*(\d+(?:\.\d{3})*,\d{2})\b'
            },
            'fiscal': {
                'tipo_documento': r'\b(NOTA FISCAL|CUPOM FISCAL|NFe)\b',
                'numero_documento': r'\b(NFe|Nota Fiscal)\s*[Nº]?\s*(\d{6,12})\b',
                'cnpj': r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b',
                'data_emissao': r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b',
                'valor_total': r'\bVALOR\s*(?:TOTAL)?\s*[R$]?\s*(\d+(?:\.\d{3})*,\d{2})\b',
                'tipo_pagamento': r'\b(BOLETO|PIX|TRANSFERÊNCIA|CARTÃO)\b'
            },
            'bancario': {
                'tipo_documento': r'\b(EXTRATO|COMPROVANTE|CONTRACHEQUE)\b',
                'conta': r'\bCONTA\s*[Nº]?\s*(\d{6,10})\b',
                'agencia': r'\bAGÊNCIA\s*[Nº]?\s*(\d{4})\b',
                'data_lancamento': r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b',
                'valor_lancamento': r'\bVALOR\s*[R$]?\s*(\d+(?:\.\d{3})*,\d{2})\b',
                'tipo_lancamento': r'\b(CRÉDITO|DÉBITO|TRANSFERÊNCIA|PAGAMENTO)\b'
            }
        }
    
    def normalize_value(self, value: str) -> float:
        if not value:
            return 0.0
        value = value.replace('R$', '').replace('.', '').replace(',', '.')
        try:
            return float(value)
        except ValueError:
            return 0.0
    
    def normalize_date(self, date_str: str) -> Optional[datetime]:
        if not date_str:
            return None
        date_formats = [
            '%d/%m/%Y',
            '%d/%m/%y',
            '%m/%d/%Y',
            '%Y-%m-%d'
        ]
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue
        return None
    
    def extract_document_data(self, text: str, documento_tipo: Optional[str] = None) -> Dict[str, Any]:
        if not documento_tipo:
            for tipo, patterns in self.extraction_patterns.items():
                if re.search(patterns['tipo_documento'], text, re.IGNORECASE):
                    documento_tipo = tipo
                    break
        
        if not documento_tipo:
            return {
                'tipo_documento': 'Não identificado',
                'texto_original': text[:500]
            }
        
        dados_extraidos = {
            'tipo_documento': documento_tipo.capitalize()
        }
        
        patterns = self.extraction_patterns.get(documento_tipo, {})
        
        for campo, padrao in patterns.items():
            if campo == 'tipo_documento':
                continue
            
            match = re.search(padrao, text, re.IGNORECASE)
            if match:
                valor = match.group(1) if match.groups() else match.group(0)
                
                if 'valor' in campo or 'total' in campo:
                    dados_extraidos[campo] = self.normalize_value(valor)
                elif 'data' in campo:
                    dados_extraidos[campo] = self.normalize_date(valor)
                else:
                    dados_extraidos[campo] = valor
        
        if self.mistral_client:
            dados_extraidos = self._enrich_with_ai(text, dados_extraidos)
        
        return dados_extraidos
    
    def _enrich_with_ai(self, text: str, dados_extraidos: Dict[str, Any]) -> Dict[str, Any]:
        try:
            messages = [
                {
                    "role": "system",
                    "content": "Você é um assistente especializado em extração de informações de documentos. Analise o texto fornecido e extraia informações adicionais não capturadas pelos padrões básicos. Forneça dados em formato JSON."
                },
                {
                    "role": "user",
                    "content": f"Dados já extraídos: {json.dumps(dados_extraidos)}\n\nTexto do documento: {text[:2000]}\n\nPor favor, forneça informações adicionais relevantes em JSON. Foque em campos não preenchidos que possam ser importantes."
                }
            ]
            
            response = self.mistral_client.chat.complete(
                model="mistral-large-latest",
                messages=messages,
                response_format={"type": "json_object"}
            )
            
            dados_ai = json.loads(response.choices[0].message.content)
            dados_extraidos.update(dados_ai)
        
        except Exception as e:
            logging.error(f"Erro no enriquecimento por IA: {e}")
        
        return dados_extraidos
    
    def process_document_batch(self, documentos: List[str]) -> pd.DataFrame:
        dados_documentos = []
        
        for documento in documentos:
            try:
                texto = extract_text(documento)
                dados = self.extract_document_data(texto)
                dados['caminho_documento'] = documento
                dados_documentos.append(dados)
            
            except Exception as e:
                logging.error(f"Erro ao processar documento {documento}: {e}")
        
        return pd.DataFrame(dados_documentos)

