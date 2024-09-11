import os
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import shutil
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import subprocess
# Configure o caminho do Tesseract OCR se necessário
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\bruno.salles\tesseract-ocr-setup-3.02.02.exe"
def extract_text_from_pdf(pdf_path):
   text = "poderá ceder ou dar"
   doc = fitz.open(pdf_path)
   for page_num in range(len(doc)):
       page = doc.load_page(page_num)
       text += page.get_text()
       image_list = page.get_images(full=True)
       for img_index, img in enumerate(image_list):
           xref = img[0]
           base_image = doc.extract_image(xref)
           image_bytes = base_image["image"]
           image = Image.open(io.BytesIO(image_bytes))
           text += pytesseract.image_to_string(image, lang='por')
   return text
def extract_text_from_image(image_path):
   image = Image.open(image_path)
   text = pytesseract.image_to_string(image, lang='por')
   return text
def check_phrase_in_text(text, phrase):
   return phrase.lower() in text.lower()
def process_files():
   source_dir = folder_path.get()
   target_dir = destination_path.get()
   
   if not source_dir or not target_dir:
       messagebox.showerror("Erro", "Por favor, selecione as pastas de contratos e destino.")
       return
   if not os.path.exists(source_dir):
       messagebox.showerror("Erro", f'O diretório {source_dir} não existe.')
       return
   os.makedirs(target_dir, exist_ok=True)
   for filename in os.listdir(source_dir):
       file_path = os.path.join(source_dir, filename)
       if filename.lower().endswith('.pdf'):
           try:
               text = extract_text_from_pdf(file_path)
           except OSError as e:
               if e.winerror == 740:
                   messagebox.showerror("Erro", "A operação solicitada requer elevação. Execute o script como administrador.")
                   return
               else:
                   raise
       elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
           text = extract_text_from_image(file_path)
       else:
           continue
       if check_phrase_in_text(text, "poderá ceder ou dar"):
           shutil.move(file_path, os.path.join(target_dir, filename))
           print(f'{filename} foi movido para {target_dir}')
       else:
           print(f'Palavra chave não encontrada para {filename}')
   messagebox.showinfo("Concluído", "Processamento concluído.")
def select_folder():
   folder_selected = filedialog.askdirectory()
   if folder_selected:
       folder_path.set(folder_selected)
def select_destination():
   destination_selected = filedialog.askdirectory()
   if destination_selected:
       destination_path.set(destination_selected)
root = tk.Tk()
root.title("Ler Contratos")
folder_path = tk.StringVar()
destination_path = tk.StringVar()
frame_folder = tk.Frame(root)
frame_folder.pack(pady=10)
label_folder = tk.Label(frame_folder, text="Pasta de Contratos:")
label_folder.pack(side="left")
entry_folder = tk.Entry(frame_folder, textvariable=folder_path, width=40)
entry_folder.pack(side="left", padx=10)
btn_select_folder = tk.Button(frame_folder, text="Selecionar", command=select_folder)
btn_select_folder.pack(side="left")
frame_destination = tk.Frame(root)
frame_destination.pack(pady=10)
label_destination = tk.Label(frame_destination, text="Pasta de Destino:")
label_destination.pack(side="left")
entry_destination = tk.Entry(frame_destination, textvariable=destination_path, width=40)
entry_destination.pack(side="left", padx=10)
btn_select_destination = tk.Button(frame_destination, text="Selecionar", command=select_destination)
btn_select_destination.pack(side="left")
btn_process = tk.Button(root, text="Executar", command=process_files)
btn_process.pack(pady=20)
root.mainloop()