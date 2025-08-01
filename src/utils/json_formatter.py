

import json
import re
from typing import Dict, List, Tuple, Optional

class JsonFormatter:
    @staticmethod
    def create_mistral_entry(text: str, paragraphs: List[Tuple[str, str]]) -> Optional[Dict]:
        """Cria entrada garantindo que termina com 'assistant'"""
        messages = []
        
        messages.append({
            "role": "user",
            "content": JsonFormatter.sanitize_text(text[:5000])
        })
        
        assistant_content = "\n\n".join([p[0] for p in paragraphs if p[0].strip()])
        
        if not assistant_content:
            return None
            
        messages.append({
            "role": "assistant",
            "content": JsonFormatter.sanitize_text(assistant_content)
        })
        
        return {"messages": messages} if len(messages) >= 2 else None

    @staticmethod
    def sanitize_text(text: str) -> str:
        """Sanitização mais rigorosa para compatibilidade com LLMs"""
        cleaned = re.sub(r'\s+', ' ', text.strip())
        return cleaned.encode('ascii', 'ignore').decode()

