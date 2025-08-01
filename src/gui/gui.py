import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import logging
import os
from ..ocr.tesseract_ocr import TesseractOCR
from ..ocr.mistral_ocr import MistralOCR
from ..utils.docx_formatter import DocxFormatter
from ..utils.json_formatter import JsonFormatter
from ..utils.document_data_extractor import DocumentDataExtractor
from ..utils.document_enhancer import DocumentEnhancer
import threading
import datetime

class PDFProcessorApp(tk.Tk):
    """Interface gráfica principal"""

    def __init__(self):
        super().__init__()
        self.title("OCR Processor Pro v5 - Tesseract & Mistral OCR")
        self.geometry("800x700")
        
        self.tesseract_ocr = TesseractOCR()
        self.mistral_ocr = MistralOCR()
        self.current_ocr = self.tesseract_ocr
        
        self._json_write_lock = threading.Lock()
        
        self._setup_ui()
        self._update_api_stats()

    def _setup_ui(self):
        self.input_dir_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.lang_var = tk.StringVar(value='por')
        self.ocr_type_var = tk.StringVar(value='tesseract')
        
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side='left', fill='both', expand=True)
        
        right_frame = ttk.LabelFrame(main_frame, text="Log de Operações")
        right_frame.pack(side='right', fill='both', expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(right_frame, state='disabled', width=40, height=30)
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        logging.basicConfig(level=logging.INFO,
                        handlers=[self._create_file_handler(),
                                self._create_gui_handler(self.log_text)])
        
        self._create_api_status_frame(left_frame)
        self._create_directory_selector(left_frame)
        self._create_ocr_type_selector(left_frame)
        self._create_language_selector(left_frame)
        
        self.generate_summary_var = tk.BooleanVar(value=False)
        summary_frame = ttk.Frame(left_frame)
        summary_frame.pack(fill='x', padx=10, pady=5)
        ttk.Checkbutton(
            summary_frame, 
            text="Gerar sumário e tabela de conteúdo usando IA", 
            variable=self.generate_summary_var
        ).pack(anchor='w')
        
        self.extract_data_var = tk.BooleanVar(value=False)
        extract_frame = ttk.Frame(left_frame)
        extract_frame.pack(fill='x', padx=10, pady=5)
        ttk.Checkbutton(
            extract_frame, 
            text="Extrair dados estruturados dos documentos", 
            variable=self.extract_data_var
        ).pack(anchor='w')
        
        self._create_controls(left_frame)
        self._create_progress_bar(left_frame)

    def _create_api_status_frame(self, parent_frame):
        status_frame = ttk.LabelFrame(parent_frame, text="Status da API")
        status_frame.pack(fill='x', padx=10, pady=5)
        status_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(status_frame, text="Conexão:").grid(row=0, column=0, sticky='w')
        self.api_connection_label = ttk.Label(status_frame, text="Desconectado", foreground="red")
        self.api_connection_label.grid(row=0, column=1, sticky='w')

    def _create_directory_selector(self, parent_frame):
        dir_frame = ttk.Frame(parent_frame)
        dir_frame.pack(fill='x', padx=10, pady=5)

        for label, var in [("Entrada:", self.input_dir_var),
                        ("Saída:", self.output_dir_var)]:
            frame = ttk.Frame(dir_frame)
            frame.pack(fill='x', pady=2)

            ttk.Label(frame, text=label).pack(side='left')
            ttk.Entry(frame, textvariable=var, width=40).pack(side='left', expand=True)
            ttk.Button(frame, text="📁", command=lambda v=var: v.set(filedialog.askdirectory()))\
                .pack(side='left')

    def _create_ocr_type_selector(self, parent_frame):
        ocr_frame = ttk.LabelFrame(parent_frame, text="Tipo de OCR")
        ocr_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Radiobutton(
            ocr_frame, 
            text="Tesseract OCR (local)", 
            variable=self.ocr_type_var,
            value="tesseract",
            command=self._update_ocr_processor
        ).pack(anchor='w', padx=10)
        
        ttk.Radiobutton(
            ocr_frame, 
            text="Mistral OCR (digitar API Key)", 
            variable=self.ocr_type_var,
            value="mistral",
            command=self._update_ocr_processor
        ).pack(anchor='w', padx=10)
        
        ttk.Radiobutton(
            ocr_frame, 
            text="Mistral OCR (usar API Key em arquivo)", 
            variable=self.ocr_type_var,
            value="mistral_file",
            command=self._update_ocr_processor
        ).pack(anchor='w', padx=10)
        
        self.mistral_config_frame = ttk.Frame(ocr_frame)
        self.mistral_config_frame.pack(fill='x', pady=5)
        
    def _create_language_selector(self, parent_frame):
        lang_frame = ttk.Frame(parent_frame)
        lang_frame.pack(pady=5)

        ttk.Label(lang_frame, text="Idioma:").pack(side='left')
        ttk.Combobox(lang_frame, textvariable=self.lang_var,
                values=['por', 'eng', 'spa', 'fra', 'deu'], state='readonly').pack(side='left')

    def _create_controls(self, parent_frame):
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Iniciar", command=self._start_processing).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=self._cancel_processing).pack(side='left', padx=5)
        
    def _create_progress_bar(self, parent_frame):
        progress_frame = ttk.Frame(parent_frame)
        progress_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(progress_frame, text="Progresso:").pack(side='left')
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, length=300)
        self.progress_bar.pack(side='left', fill='x', expand=True, padx=5)
        
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.pack(side='left')
        ttk.Label(progress_frame, text="%").pack(side='left')
        
    def _update_api_key(self):
        new_key = self.api_key_entry.get().strip()
        if new_key:
            self.mistral_ocr.api_key = new_key
            self.api_status_label.config(
                text="Status: API key configurada (não testada)",
                foreground="blue"
            )
            messagebox.showinfo("Sucesso", "API Key atualizada com sucesso! Recomendamos testar a conexão.")
        else:
            messagebox.showwarning("Aviso", "API Key não pode estar vazia.")
            self.api_status_label.config(
                text="Status: API key não configurada",
                foreground="orange"
            )
    
    def _test_mistral_api(self):
        api_key = self.api_key_entry.get().strip()
        
        if not api_key:
            messagebox.showwarning("Aviso", "Insira uma API Key para testar a conexão!")
            return
        
        self.api_status_label.config(
            text="Status: Testando conexão...",
            foreground="blue"
        )
        self.update_idletasks()
        
        try:
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
            
            test_payload = {
                "model": "mistral-ocr-latest",
                "id": str(uuid.uuid4()),
                "document": {
                    "type": "document_base64",
                    "document_base64": "SGVsbG8gV29ybGQ=",
                    "document_name": "test.txt"
                }
            }
            
            response = requests.post(
                "https://api.mistral.ai/v1/ocr",
                headers=headers,
                json=test_payload,
                timeout=10
            )
            
            if response.status_code == 200:
                self.api_status_label.config(
                    text="Status: Conexão bem-sucedida! API pronta para uso.",
                    foreground="green"
                )
                messagebox.showinfo("Sucesso", "Conexão com a API Mistral OCR estabelecida com sucesso!")
            elif response.status_code == 401:
                self.api_status_label.config(
                    text="Status: Falha de autenticação - API Key inválida!",
                    foreground="red"
                )
                messagebox.showerror("Erro", "API Key inválida ou expirada. Verifique suas credenciais.")
            elif response.status_code == 422:
                self.api_status_label.config(
                    text="Status: API Key válida, mas requisição de teste precisa de ajustes.",
                    foreground="green"
                )
                messagebox.showinfo("Parcialmente Sucesso", "API Key aceita, mas a estrutura da requisição precisa de ajustes.")
            else:
                self.api_status_label.config(
                    text=f"Status: Erro na conexão - Código {response.status_code}",
                    foreground="red"
                )
                messagebox.showerror("Erro", f"Erro ao conectar com a API Mistral OCR. Código: {response.status_code}")
                
        except requests.exceptions.ConnectionError:
            self.api_status_label.config(
                text="Status: Falha de conexão - Verifique sua internet",
                foreground="red"
            )
            messagebox.showerror("Erro", "Não foi possível conectar ao servidor da API. Verifique sua conexão de internet.")
        except requests.exceptions.Timeout:
            self.api_status_label.config(
                text="Status: Timeout na conexão com a API",
                foreground="red"
            )
            messagebox.showerror("Erro", "Tempo de conexão esgotado. O servidor pode estar sobrecarregado.")
        except Exception as e:
            self.api_status_label.config(
                text=f"Status: Erro desconhecido na conexão",
                foreground="red"
            )
            messagebox.showerror("Erro", f"Erro ao testar conexão: {str(e)}")
    
    def _update_ocr_processor(self):
        ocr_type = self.ocr_type_var.get()
        
        if ocr_type == "tesseract":
            self.current_ocr = self.tesseract_ocr
            logging.info("Usando Tesseract OCR para processamento")
            self.mistral_config_frame.pack_forget()
            
        elif ocr_type == "mistral":
            self.mistral_ocr.api_key = self.api_key_entry.get().strip()
            self.current_ocr = self.mistral_ocr
            logging.info("Usando Mistral OCR para processamento")
            self.mistral_config_frame.pack(fill='x', pady=5)
            
        elif ocr_type == "mistral_file":
            api_key = self._load_api_key_from_file()
            if api_key:
                self.mistral_ocr.api_key = api_key
                self.api_key_entry.delete(0, tk.END)
                self.api_key_entry.insert(0, "********" + api_key[-4:])
                self.current_ocr = self.mistral_ocr
                logging.info("Usando Mistral OCR com chave carregada do arquivo")
                self.mistral_config_frame.pack(fill='x', pady=5)
                self.api_status_label.config(
                    text="Status: API key carregada do arquivo",
                    foreground="green"
                )
            else:
                messagebox.showwarning("Aviso", "Não foi possível carregar a chave da API do arquivo. Por favor, digite uma chave manualmente.")
                self.ocr_type_var.set("mistral")
                self.mistral_config_frame.pack(fill='x', pady=5)
                self.api_status_label.config(
                    text="Status: Falha ao carregar chave do arquivo",
                    foreground="red"
                )
        
        if ocr_type in ["mistral", "mistral_file"] and self.mistral_ocr.api_key:
            self.api_status_label.config(
                text="Status: API key configurada, pronta para processamento",
                foreground="green"
            )
        elif ocr_type in ["mistral", "mistral_file"]:
            self.api_status_label.config(
                text="Status: API key não configurada",
                foreground="orange"
            )
      
    def _cancel_processing(self):
        self.current_ocr.stop_event.set()
        logging.info("Processamento cancelado pelo usuário")

    def _create_file_handler(self):
        return logging.handlers.RotatingFileHandler('processing.log', maxBytes=5 * 1024 * 1024,
                                 backupCount=5, encoding='utf-8')

    def _create_gui_handler(self, widget):
        class GuiHandler(logging.Handler):
            def emit(self, record):
                widget.configure(state='normal')
                widget.insert(tk.END, self.format(record) + '\n')
                widget.see(tk.END)
                widget.configure(state='disabled')
        
        handler = GuiHandler()
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        return handler

    def _start_processing(self):
        input_dir = self.input_dir_var.get()
        output_dir = self.output_dir_var.get()

        if not (input_dir and output_dir):
            messagebox.showwarning("Aviso", "Selecione os diretórios de entrada e saída!")
            return
        
        ocr_type = self.ocr_type_var.get()
        
        if ocr_type == "mistral":
            if not self.api_key_entry.get().strip():
                messagebox.showwarning("Aviso", "API Key do Mistral OCR não configurada. Por favor, configure uma chave válida.")
                return        
        elif ocr_type == "tesseract":
            pass

        self.tesseract_ocr.stop_event.clear()
        self.mistral_ocr.stop_event.clear()
        
        self._update_ocr_processor()
        
        self.progress_var.set(0)
        logging.info(f"Iniciando processamento com {ocr_type.upper()} OCR")
        
        processing_thread = threading.Thread(
            target=self._process_files,
            args=(input_dir, output_dir),
            daemon=True
        )
        processing_thread.start()

    def _process_files(self, input_dir: str, output_dir: str):
        try:
            pdfs = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
            images = [f for f in os.listdir(input_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp'))]
            
            all_files = pdfs + images
            
            if not all_files:
                messagebox.showinfo("Informação", "Nenhum arquivo PDF ou imagem encontrado no diretório de entrada.")
                return
            
            logging.info(f"Iniciando processamento de {len(all_files)} arquivos ({len(pdfs)} PDFs, {len(images)} imagens) com {self.ocr_type_var.get()} OCR")

            processed_files = []

            with threading.ThreadPoolExecutor() as executor:
                futures = {executor.submit(self._process_single_file_or_image,
                                        os.path.join(input_dir, f),
                                        output_dir): f for f in all_files}

                for i, future in enumerate(as_completed(futures), 1):
                    file_name = futures[future]
                    result = future.result()
                    processed_files.append((os.path.join(input_dir, file_name), result))
                    
                    progress = (i / len(all_files)) * 100
                    self.progress_var.set(round(progress, 1))
                    self.update_idletasks()
                    
                    if self.current_ocr.stop_event.is_set():
                        break

            if self.extract_data_var.get() and hasattr(self, 'mistral_ocr') and self.mistral_ocr.api_key:
                logging.info("Iniciando extração de dados estruturados...")
                
                successful_files = [path for path, success in processed_files if success]
                
                if successful_files:
                    try:
                        extractor = DocumentDataExtractor(self.mistral_ocr.api_key)
                        data_df = extractor.process_document_batch(successful_files)
                        csv_path = os.path.join(output_dir, "dados_extraidos.csv")
                        data_df.to_csv(csv_path, index=False)
                        logging.info(f"Dados estruturados salvos em {csv_path}")
                        messagebox.showinfo("Extração de Dados", f"Dados estruturados extraídos de {len(successful_files)} documentos e salvos em {csv_path}")
                    except Exception as e:
                        logging.error(f"Erro na extração de dados: {e}")
                        messagebox.showerror("Erro", f"Falha na extração de dados estruturados: {e}")
                else:
                    logging.warning("Nenhum arquivo processado com sucesso para extração de dados.")

            if not self.current_ocr.stop_event.is_set():
                messagebox.showinfo("Concluído", f"Processamento concluído com sucesso! {i} de {len(all_files)} arquivos processados.")
            else:
                messagebox.showinfo("Interrompido", f"Operação interrompida. {i} de {len(all_files)} arquivos processados.")

        except Exception as e:
            logging.error(f"Erro crítico: {e}")
            messagebox.showerror("Erro", f"Falha no processamento: {e}")

    def _process_single_file_or_image(self, file_path: str, output_dir: str) -> bool:
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext in ['.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp']:
            return self._process_single_image(file_path, output_dir)
        else:
            return self._process_single_file(file_path, output_dir)
    
    def _process_single_image(self, image_path: str, output_dir: str) -> bool:
        try:
            file_name = os.path.basename(image_path)
            logging.info(f"Processando imagem {file_name}")
            
            with Image.open(image_path) as img:
                preprocessed = self.tesseract_ocr._preprocess_image(img)
                text = pytesseract.image_to_string(preprocessed, lang=self.lang_var.get())
                
                if not text or len(text.strip()) < 50:
                    logging.warning(f"Texto extraído da imagem {file_name} é muito curto ou vazio.")
                    return False
                
                paragraphs = self.current_ocr.get_paragraphs(text)
                
                docx_path = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}.docx")
                
                doc = Document()
                DocxFormatter.setup_document_styles(doc)
                doc.add_heading(f"Imagem: {file_name}", level=0)
                doc.add_paragraph(f"Processado com: Tesseract OCR")
                doc.add_paragraph(f"Data de processamento: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
                doc.add_paragraph(f"Idioma: {self.lang_var.get()}")
                doc.add_paragraph("").paragraph_format.space_after = Pt(20)
                
                for para_text, para_type in paragraphs:
                    DocxFormatter.add_paragraph_with_style(doc, para_text, para_type)
                
                doc.save(docx_path)
                logging.info(f"Documento DOCX criado para imagem: {docx_path}")
                
                self._generate_json(image_path, output_dir, text, paragraphs)
                
                if self.generate_summary_var.get():
                    self._generate_summary_and_toc(docx_path, text)
                
                return True
                
        except Exception as e:
            logging.error(f"Erro ao processar imagem {image_path}: {e}")
            return False
    
    def _process_single_file(self, file_path: str, output_dir: str):
        try:
            file_name = os.path.basename(file_path)
            logging.info(f"Processando {file_name} com {self.ocr_type_var.get()} OCR")
            
            text = self.current_ocr.extract_text(file_path, self.lang_var.get())
            
            if not text or text.startswith("Erro:"):
                logging.error(f"Falha ao extrair texto de {file_name}: {text}")
                return False
            
            paragraphs = self.current_ocr.get_paragraphs(text)
            
            docx_success = self._generate_docx(file_path, output_dir, paragraphs)
            json_success = self._generate_json(file_path, output_dir, text, paragraphs)
            
            if self.generate_summary_var.get() and docx_success:
                docx_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}.docx")
                self._generate_summary_and_toc(docx_path, text)
            
            return docx_success and json_success
            
        except Exception as e:
            logging.error(f"Erro desconhecido ao processar {file_path}: {e}")
            return False
    
    def _generate_summary_and_toc(self, docx_path: str, text: str) -> bool:
        try:
            if not hasattr(self, 'mistral_ocr') or not self.mistral_ocr.api_key:
                logging.warning("API Key do Mistral não configurada. Não é possível gerar sumário.")
                messagebox.showwarning("Aviso", "API Key do Mistral é necessária para gerar sumário. Por favor, configure uma chave válida.")
                return False
            
            logging.info(f"Iniciando geração de sumário e tabela de conteúdo para {os.path.basename(docx_path)}")
            
            enhancer = DocumentEnhancer(self.mistral_ocr.api_key)
            success = enhancer.process_document(docx_path, "mistral-large-latest")
            
            if success:
                logging.info(f"Sumário e tabela de conteúdo gerados com sucesso para {os.path.basename(docx_path)}")
                return True
            else:
                logging.error(f"Falha ao gerar sumário e tabela de conteúdo para {os.path.basename(docx_path)}")
                return False
                
        except Exception as e:
            logging.error(f"Erro ao gerar sumário e tabela de conteúdo: {e}")
            return False

    def _generate_docx(self, file_path: str, output_dir: str, paragraphs: List[Tuple[str, str]]):
        try:
            docx_path = os.path.join(output_dir,
                              f"{os.path.splitext(os.path.basename(file_path))[0]}.docx")
            
            doc = Document()
            DocxFormatter.setup_document_styles(doc)
            doc.add_heading(f"Documento: {os.path.basename(file_path)}", level=0)
            doc.add_paragraph(f"Processado com: {self.ocr_type_var.get().capitalize()} OCR")
            doc.add_paragraph(f"Data de processamento: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
            doc.add_paragraph(f"Idioma: {self.lang_var.get()}")
            doc.add_paragraph("").paragraph_format.space_after = Pt(20)
            
            for para_text, para_type in paragraphs:
                DocxFormatter.add_paragraph_with_style(doc, para_text, para_type)
            
            doc.save(docx_path)
            logging.info(f"Documento DOCX criado: {docx_path}")
            
        except Exception as e:
            logging.error(f"Erro ao gerar DOCX para {file_path}: {e}")
            return False
        
        return True
    
    def _generate_json(self, file_path: str, output_dir: str, text: str, paragraphs: List[Tuple[str, str]]):
        try:
            entry = JsonFormatter.create_mistral_entry(text, paragraphs)
            
            if not entry:
                logging.warning(f"Ignorando entrada inválida para {file_path}")
                return False
                
            output_file = os.path.join(output_dir, "mistral_dataset.jsonl")
            
            with self._json_write_lock:
                with open(output_file, 'a', encoding='utf-8') as f:
                    f.write(json.dumps(entry, ensure_ascii=False) + '\n')
                
            return True
        except Exception as e:
            logging.error(f"Erro JSON: {e}")
            return False
        
    def _update_api_stats(self):
        if hasattr(self, 'active_requests_label') and hasattr(self, 'mistral_ocr'):
            self.active_requests_label.config(text=str(self.mistral_ocr.active_requests))
        
        if hasattr(self, 'tokens_used_label') and hasattr(self, 'mistral_ocr'):
            self.tokens_used_label.config(text=str(self.mistral_ocr.total_tokens_used))
        
        if hasattr(self, 'total_calls_label') and hasattr(self, 'mistral_ocr'):
            self.total_calls_label.config(text=str(self.mistral_ocr.api_calls_count))
        
        if hasattr(self, 'api_connection_label') and hasattr(self, 'mistral_ocr'):
            if self.mistral_ocr.api_key:
                self.api_connection_label.config(text="Conectado", foreground="green")
            else:
                self.api_connection_label.config(text="Desconectado", foreground="red")
        
        self.after(2000, self._update_api_stats)