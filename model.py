import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import pdfplumber
import re
import pandas as pd
import sqlite3
import os
# Configurar o caminho para o executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'Tesseract-OCR\tesseract.exe'
def preprocess_image(image):
   image = image.convert('L')
   image = image.filter(ImageFilter.SHARPEN)
   enhancer = ImageEnhance.Contrast(image)
   image = enhancer.enhance(4)
   return image
def extract_text_from_image(image):
   processed_img = preprocess_image(image)
   text = pytesseract.image_to_string(processed_img, lang='por', config='--psm 6')
   text = re.sub(r'[^\w\s.,-]', '', text)
   text = re.sub(r'(\d+),(\d{2})', r'\1.\2', text)
   text = re.sub(r'\s+', ' ', text).strip()
   return text
def search_words_in_pdf(pdf_path, words_to_search):
   found_words = {word.lower(): False for word in words_to_search}
   try:
       with pdfplumber.open(pdf_path) as pdf:
           for page in pdf.pages:
               text = page.extract_text()
               if not text:
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