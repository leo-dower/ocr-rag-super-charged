# ocr-rag-super-charged
ocr and rag 
# Dependencies and Setup Guide for PDF Processor

This PDF processing application requires several dependencies to run properly. Below is a comprehensive guide to install all necessary components on a new PC.

## Required Components

### Core Requirements:
1. **Python 3.6+**: The base programming language
2. **Tesseract OCR**: Open-source OCR engine
3. **Poppler**: PDF rendering library
4. **Python packages**:
   - python-docx (for creating Word documents)
   - beautifulsoup4 (for HTML parsing)
   - pytesseract (Python wrapper for Tesseract)
   - pdf2image (converts PDFs to images)
   - Pillow (image processing)
   - pdfminer.six (PDF text extraction)
   - pydantic (data validation)
   - pandas (data manipulation)
   - mistralai (API client for Mistral AI)

## Step-by-Step Installation Guide

### 1. Install Python

**Windows**:
1. Download the Python installer from [python.org](https://www.python.org/downloads/)
2. Run the installer and check the box that says "Add Python to PATH"
3. Complete the installation

**macOS**:
```bash
brew install python
```

**Linux (Ubuntu/Debian)**:
```bash
sudo apt update
sudo apt install python3 python3-pip
```

### 2. Install Tesseract OCR

**Windows**:
1. Download the installer from [UB-Mannheim's GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
2. Run the installer
3. Ensure the installation path is added to the system PATH
   - Default location: `C:\Program Files\Tesseract-OCR`

**macOS**:
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian)**:
```bash
sudo apt update
sudo apt install tesseract-ocr
sudo apt install libtesseract-dev
```

### 3. Install Poppler

**Windows**:
1. Download the Windows build from [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases/)
2. Extract the files to a folder (e.g., `C:\Poppler`)
3. Add the bin directory to PATH:
   - Open Control Panel > System > Advanced System Settings
   - Click "Environment Variables"
   - In "System Variables", select "Path" and click "Edit"
   - Add the path to the bin directory (e.g., `C:\Poppler\bin`)

**macOS**:
```bash
brew install poppler
```

**Linux (Ubuntu/Debian)**:
```bash
sudo apt update
sudo apt install poppler-utils
```

### 4. Install Required Python Packages

The script will attempt to install missing packages automatically, but you can install them all at once with:

```bash
pip install python-docx beautifulsoup4 pytesseract pdf2image Pillow pdfminer.six pydantic pandas mistralai
```

### 5. Language Data for Tesseract (Optional)

If you need languages other than English:

**Windows**:
- The installer should offer language selection during installation

**macOS/Linux**:
```bash
# For Portuguese
sudo apt install tesseract-ocr-por

# For Spanish
sudo apt install tesseract-ocr-spa

# For French
sudo apt install tesseract-ocr-fra

# For German
sudo apt install tesseract-ocr-deu
```

## Running the Application

1. Save the Python script to your computer (e.g., `pdf_processor.py`)
2. Open a terminal/command prompt
3. Navigate to the folder containing the script
4. Run the script:
   ```bash
   python pdf_processor.py
   ```

## First-Run Configuration

When you first run the application:

1. The GUI will open
2. Select input and output directories using the folder icon
3. Choose OCR type (Tesseract for local processing, Mistral for API-based processing)
4. If using Mistral OCR, you'll need to enter an API key
5. Select the language for OCR
6. Click "Start" to begin processing

## Troubleshooting Common Issues

- **"Tesseract not found" error**: Ensure Tesseract is installed and added to your PATH
- **"Poppler not found" warning**: Ensure Poppler is installed and added to your PATH
- **PDF processing fails**: Make sure both Tesseract and Poppler are correctly installed
- **Missing Python packages**: The script should install them automatically, but you can install them manually using pip

The script includes comprehensive error handling, so it should provide helpful error messages when issues arise. Pay attention to the log window in the application for specific error details.

This application combines multiple technologies for advanced document processing, so proper setup of all components is essential for full functionality.
