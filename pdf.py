import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
def pdf_to_word(pdf_file, word_file):
   try:
       cv = Converter(pdf_file)
       cv.convert(word_file, start=0, end=None)
       cv.close()
       messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso: {word_file}")
   except Exception as e:
       messagebox.showerror("Erro", f"Erro ao converter o arquivo: {str(e)}")
def select_pdf():
   pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
   if pdf_file:
       pdf_entry.delete(0, tk.END)
       pdf_entry.insert(0, pdf_file)
def select_word():
   word_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
   if word_file:
       word_entry.delete(0, tk.END)
       word_entry.insert(0, word_file)
def convert():
   pdf_file = pdf_entry.get()
   word_file = word_entry.get()
   if pdf_file and word_file:
       pdf_to_word(pdf_file, word_file)
   else:
       messagebox.showwarning("Aviso", "Por favor, selecione um arquivo PDF e um local para salvar o arquivo Word.")
app = tk.Tk()
app.title("Conversor de PDF para Word")
tk.Label(app, text="Selecionar arquivo PDF:").grid(row=0, column=0, padx=10, pady=10)
pdf_entry = tk.Entry(app, width=50)
pdf_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Procurar", command=select_pdf).grid(row=0, column=2, padx=10, pady=10)
tk.Label(app, text="Salvar como arquivo Word:").grid(row=1, column=0, padx=10, pady=10)
word_entry = tk.Entry(app, width=50)
word_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Procurar", command=select_word).grid(row=1, column=2, padx=10, pady=10)
tk.Button(app, text="Converter", command=convert).grid(row=2, column=1, pady=20)
app.mainloop()