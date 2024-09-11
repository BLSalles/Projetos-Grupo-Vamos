import pytesseract
from PIL import Image, ImageEnhance, ImageFilter, ImageTk  # Biblioteca para trabalhar com imagens
import pdfplumber
import os
import re
import pandas as pd
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import traceback
# Configurar o caminho para o executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'Tesseract-OCR\tesseract.exe'
def preprocess_image(image):
   # Converter para escala de cinza
   image = image.convert('L')
   # Aplicar filtro de nitidez
   image = image.filter(ImageFilter.SHARPEN)
   # Aumentar o contraste
   enhancer = ImageEnhance.Contrast(image)
   image = enhancer.enhance(4)
   return image
def extract_text_from_image(image):
   # Aplicar pré-processamento
   processed_img = preprocess_image(image)
   # Extrair texto usando Tesseract
   text = pytesseract.image_to_string(processed_img, lang='por', config='--psm 6')
   # Tratar caracteres especiais e números
   text = re.sub(r'[^\w\s.,-]', '', text)  # Remove caracteres especiais, mantém alfanuméricos, espaço, ponto, vírgula e hífen
   text = re.sub(r'(\d+),(\d{2})', r'\1.\2', text)  # Substitui vírgula por ponto em números com duas casas decimais
   text = re.sub(r'\s+', ' ', text).strip()  # Substitui múltiplos espaços por um único espaço e remove espaços no início e no final
   return text
def search_words_in_pdf(pdf_path, words_to_search):
   found_words = {word.lower(): False for word in words_to_search}
   try:
       with pdfplumber.open(pdf_path) as pdf:
           for page in pdf.pages:
               text = page.extract_text()
               if not text:  # Se o texto não for extraído, realizar OCR
                   # Obter a imagem da página como um objeto PIL Image
                   pil_image = page.to_image().original
                   text = extract_text_from_image(pil_image)
               text = text.lower()
               for word in found_words.keys():
                   if word in text:
                       found_words[word] = True
   except Exception as e:
       print(f"Erro ao processar o PDF {pdf_path}: {e}")
   return found_words
def percentage_found(found_words):
   total_words = len(found_words)
   found_count = sum(found_words.values())
   return (found_count / total_words) * 100
def format_observacao(text):
   if not isinstance(text, str):
       text = str(text)
   text = re.sub(r'(\d+,\d{2})(,)', r'\1 \2', text)
   text = re.sub(r'(?<!\d)([.,:;])(?!\d)', r' \1', text)
   
   return text
def process_contracts(excel_path, contracts_folder):
   df = pd.read_excel(excel_path, dtype={'Observação': str})
   conn = sqlite3.connect('status.db')
   cursor = conn.cursor()
   cursor.execute('''CREATE TABLE IF NOT EXISTS status (
                         contrato TEXT,
                         data TEXT,
                         hora TEXT,
                         percentual REAL,
                         status TEXT
                     )''')
   for index, row in df.iterrows():
       contrato_num = str(row['Documento/"simulação"']).split('-')[0]
       observacao = format_observacao(row['Observação'])
       pdf_path = None
       for file in os.listdir(contracts_folder):
           if contrato_num in file and file.endswith('.pdf'):
               pdf_path = os.path.join(contracts_folder, file)
               break
       if not pdf_path:
           df.at[index, 'Percentual(%)'] = f'{0.0}%'
           df.at[index, 'Status'] = f"Contrato {contrato_num} não encontrado."
           continue
       words_to_search = set(word.lower() for word in observacao.replace("\n", " ").split())
       found_words = search_words_in_pdf(pdf_path, words_to_search)
       found_percentage = percentage_found(found_words)
       not_found_words = [word for word, found in found_words.items() if not found]
       not_found_words_str = ', '.join(not_found_words)
       df.at[index, 'Percentual(%)'] = f"{found_percentage:.1f}%"
       df.at[index, 'Status'] = f"das palavras foram encontradas no PDF. Palavras não encontradas: {not_found_words_str}"
       print(f"Contrato: {contrato_num}, Percentual: {found_percentage:.1f}%")
       cursor.execute("INSERT INTO status (contrato, data, hora, percentual, status) VALUES (?, ?, ?, ?, ?)",
                      (contrato_num, pd.Timestamp.now().strftime('%Y-%m-%d'), pd.Timestamp.now().strftime('%H:%M:%S'),
                       found_percentage, df.at[index, 'Status']))
       conn.commit()
   conn.close()
   df.to_excel(excel_path, index=False)
def export_to_excel():
   conn = sqlite3.connect('status.db')
   df = pd.read_sql_query("SELECT * FROM status", conn)
   df.to_excel('status.xlsx', index=False)
   conn.close()
# Funções de interface gráfica com tkinter
def select_excel_file():
   global excel_path
   excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
   excel_file_label.config(text=f"Excel File: {excel_path}")
def select_contract_folder():
   global contracts_folder
   contracts_folder = filedialog.askdirectory()
   folder_label.config(text=f"Contracts Folder: {contracts_folder}")
def run_processing():
   try:
       process_contracts(excel_path, contracts_folder)
       messagebox.showinfo("Success", "Processamento concluído!")
   except Exception as e:
       error_message = f"Error: {str(e)}\n\n{traceback.format_exc()}"
       messagebox.showerror("Error", error_message)
def export_data():
   try:
       export_to_excel()
       messagebox.showinfo("Success", "Dados exportados para 'status.xlsx'!")
   except Exception as e:
       error_message = f"Error: {str(e)}\n\n{traceback.format_exc()}"
       messagebox.showerror("Error", error_message)
# Interface gráfica
root = tk.Tk()
root.title("Contract Processor")
root.configure(bg='#ED1B24')

# Carregando e exibindo logotipo na interface
logo_path = "Logo Vm 1.png"
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((100, 50), Image.Resampling.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(root, image=logo_photo, bg='#ED1B24')
logo_label.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='w')
excel_path = ""
contracts_folder = ""
ttk.Button(root, text="Selecionar Arquivo Excel", command=select_excel_file).grid(row=2, column=0, padx=10, pady=10)
excel_file_label = tk.Label(root, text="Nenhum arquivo selecionado", bg='#ED1B24', fg='white')
excel_file_label.grid(row=2, column=1, padx=10, pady=10)
ttk.Button(root, text="Selecionar Pasta de Contratos", command=select_contract_folder).grid(row=3, column=0, padx=10, pady=10)
folder_label = tk.Label(root, text="Nenhuma pasta selecionada", bg='#ED1B24', fg='white')
folder_label.grid(row=3, column=1, padx=10, pady=10)
ttk.Button(root, text="Executar", command=run_processing).grid(row=4, column=0, padx=10, pady=10)
ttk.Button(root, text="Exportar para Excel", command=export_data).grid(row=4, column=1, padx=10, pady=10)
root.mainloop()