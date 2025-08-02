# OCR Processor Pro

Este é um projeto de OCR (Reconhecimento Óptico de Caracteres) que unifica várias funcionalidades de processamento de documentos em uma única aplicação. Ele suporta tanto o Tesseract (para processamento local) quanto a API da Mistral (para processamento na nuvem), além de oferecer recursos avançados como extração de dados estruturados e geração de resumos com IA.

## Funcionalidades

- **Múltiplos Mecanismos de OCR:** Escolha entre o Tesseract (local) e a API da Mistral (nuvem).
- **Suporte a Vários Formatos:** Processe arquivos PDF e de imagem (JPG, PNG, etc.).
- **Formatos de Saída Flexíveis:** Salve os resultados como `.docx`, `.jsonl` (para fine-tuning), `.md` (Markdown) e `.csv` (para dados estruturados).
- **Extração de Dados com IA:** Extraia informações específicas de documentos jurídicos, fiscais e bancários.
- **Melhora de Documentos com IA:** Gere automaticamente um sumário executivo e um índice para os seus documentos.

## Instalação

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/leo-dower/ocr-rag-super-charged.git
    cd ocr-rag-super-charged
    ```

2.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Instale o Tesseract OCR:**
    -   **Windows:** Baixe e instale a partir do [instalador oficial](https://github.com/UB-Mannheim/tesseract/wiki).
    -   **Linux (Ubuntu/Debian):**
        ```bash
        sudo apt-get update
        sudo apt-get install tesseract-ocr
        ```
    -   **macOS (usando Homebrew):**
        ```bash
        brew install tesseract
        ```

## Como Usar

Para iniciar a aplicação, execute o seguinte comando na raiz do projeto:

```bash
python main.py
```

Isso abrirá a interface gráfica, onde você poderá selecionar os diretórios de entrada e saída, escolher o mecanismo de OCR e configurar outras opções.